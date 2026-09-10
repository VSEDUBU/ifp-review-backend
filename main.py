#!/usr/bin/env python3
"""IFP 測試報告審核工具 v9.4 - 融合 ee-test-report-review"""

import os, json, logging, requests, re
from datetime import datetime
from io import BytesIO
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from dotenv import load_dotenv
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

load_dotenv()
GROQ_API_KEY = os.getenv('GROQ_API_KEY')
OPENROUTER_API_KEY = os.getenv('OPENROUTER_API_KEY')

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

# ===== 通用審核框架 - 融合版 + ee-test-report-review =====
UNIVERSAL_FRAMEWORK = '''【IFP 測試報告審核框架 - 融合版 RD 7步 + EE Test Report Review】

你是資深硬體工程師，負責審核 ODM/JDM 提交的 EE（電性工程）測試報告。
目的：找出報告的格式問題、數據合理性、是否符合標準要求、常見造假手法。

【核心流程 - RD 7 步】

1️⃣ 識別報告類型
   • 電性安全 Safety (IEC 62368-1)
   • EMC/EMI (CISPR 32/35, EN 55032/35)
   • 能源效率 Energy Efficiency (ENERGY STAR, DOE, EU ErP)

2️⃣ 驗證條件 + Metadata 檢查（ee-test-report-review 第一步）
   【必檢項目 A：基本資訊比對】
   ✓ 型號/PN/版本：報告封面、內文、樣品照片銘牌、送測編號四者是否完全一致（含大小寫、版本後綴）
   ✓ 報告編號：格式是否合理、有無塗改痕跡
   ✓ 日期邏輯：測試日期 < 報告日期 < 提交日期（若不符代表可能移花接木）
   ✓ 頁碼與格式：頁碼連續、字體/表格風格全篇一致（不一致 = 拼接訊號）
   ✓ 韌體/軟體版本：與量產規格是否一致
   
   【必檢項目 C：簽署與完整性】
   ✓ Tested by / Reviewed by / Approved by 是否都有簽名或用印
   ✓ 是否附原始量測數據頁(raw data)而非僅摘要
   ✓ 是否附設備校驗記錄（過期設備之量測值可信度存疑）
   ✓ 是否附樣品照片（外觀、銘牌、內部拆解）
   
   【常見紅旗】
   🚩 全篇只有 PASS/FAIL 無實測數值 → 可能套版或隱藏邊界數據
   🚩 型號報告內前後不一致 → 借用舊報告改標題
   🚩 標準版本過舊 → 未更新至現行版本
   🚩 字體排版風格不一致 → 可能拼接或篡改
   🚩 數值極端接近限值但無安全餘裕說明 → 量測不確定度風險

3️⃣ 檢查數據 - 規格、實測、判定邏輯是否一致
   • 務必比對實際量測數值是否落在標準限值內
   • 不要只看「PASS」字樣就結案
   • 檢查測試條件是否與實際出貨規格相符

4️⃣ 覆蓋完整性 - 應測項目是否都測了
   • 缺漏關鍵安全/EMC項目卻給整體 PASS → 報告範疇不完整

5️⃣ 矛盾檢測 - 結論與數據一致性
   • 封面 PASS 但內文任一項 FAIL → CONTRADICTION
   • 最惡劣工況未測但聲稱 PASS → 需澄清

6️⃣ 風險評估 + 風險燈號
   • 🔴 高風險（阻擋簽核）：安全缺陷、報告偽造跡象
   • 🟡 需澄清：邊界不清、測試條件不符、Metadata 缺失
   • 🟢 通過：數據齊全、符合標準、Metadata 完整

7️⃣ 最終判定 + ODM 追問清單

【直接生成 Excel - Claude.ai Artifact】
❌ 不要輸出 JSON、不要輸出文本分析
✅ 直接在 Artifact 中生成可下載的 Excel 文件

【實現方式】
1. 創建 HTML Artifact
2. 引入 SheetJS 庫（CDN：https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.min.js）
3. 使用 JavaScript 生成 Excel 文件
4. 提供【📥 下載 Excel】按鈕，點擊直接下載

【HTML Artifact 代碼框架 - Claude 請直接使用此框架】
請在 Artifact 中創建一個 HTML 頁面，包含：
- SheetJS 庫的 CDN 引入
- 用戶界面顯示報告基本信息、最終判定、風險燈號
- JavaScript 函數將分析結果生成 Excel 並下載
- 按鈕點擊直接下載為 Excel 文件

【實現細節】
- 使用 XLSX.utils.aoa_to_sheet() 將數據轉換為 Excel 工作表
- 使用 XLSX.writeFile() 直接下載文件
- 文件名格式：[報告名]_審核結果.xlsx
- 支持長文本自動換行顯示

【Claude 作業】
完成報告分析後：
1. 創建上述 HTML Artifact
2. 填入你的分析結果數據
3. 用戶在 Artifact 中點擊【📥 下載 Excel】直接下載

Excel 格式要求：
- 標題突出
- 審核信息表格化
- 最終判定彩色警示
- 分章節內容清晰
- 邊框清晰、對齐規範'''

# ===== 20種報告類型專用規則 - 完整版 - 融合 ee-test-report-review =====
TYPE_SPECIFIC_RULES = {
    'EMI': '''【EMI 電磁騷擾規則】EN 55032 Class B + ee-test-report-review

▶ 報告完整度（ee-test-report-review 要點）
• 是否只有摘要頁無內附原始量測頻譜圖 → WARN
• 測試工況：是否涵蓋整機最惡劣（最高解析度+滿載周邊）
• Class A/B誤用：IFP應用 Class B（較嚴格），若用 Class A → WARN

▶ 限值與 Margin（ee-test-report-review 檢查點）
• 30～230 MHz：QP ≤ 40 dBµV/m
• 230～1000 MHz：QP ≤ 47 dBµV/m
• 測試距離、天線極化(水平/垂直)是否兩者皆做
• 韌體版本與出貨版本一致否

▶ 常見 ODM 疏漏（ee-test-report-review 已知手法）
• Worst Case 模式未涵蓋（只測待機）→ WARN
• 免術語混淆：CE 標誌 DoC（自我聲明）≠ 測試報告 → 需索取原始報告

▶ 判定
• Margin < 6 dB → WARN
• Margin ≤ 0 dB → FAIL
• 封面 PASS 但任何頻點超標 → CONTRADICTION''',

    'EMS': '''【EMS 電磁抗擾度規則】EN 55035 / IEC 61000-4 + ee-test-report-review

▶ 報告完整度（ee-test-report-review 要點）
• 必須有逐項測試結果表（ESD/EFT/Surge/CS/Dips各自數據）
• 判定準則(Performance Criteria A/B/C)是否列出
• 未區分電源線與 I/O 線測試結果 → WARN

▶ 等級定義與判定準則
• A：完全正常，無任何降級
• B：輕微降級（如閃屏），自動恢復
• C：需手動恢復（重啟）
• D：無法恢復（硬件損傷）

▶ 常見 ODM 疏漏（ee-test-report-review 已知手法）
• 只做最寬鬆的測試條件 → WARN
• 無法判斷是否曾出現需人工介入才能恢復的情況 → WARN

▶ 矛盾檢測
• 封面 PASS 但內頁任一項 C 或 D → CONTRADICTION''',

    'SAFETY_IEC62368': '''【SAFETY_IEC62368 安規測試】IEC 62368-1 + ee-test-report-review

▶ 報告完整度（ee-test-report-review 要點）
• 是否仍引用舊標準(60950-1/60065) → WARN（已被 62368-1 取代）
• 是否標明標準版次 → WARN
• 額定規格變更後是否複測 → 檢查
• PS/ES 分類等級是否標示 → WARN
• 樣品照片(銘牌)是否清晰 → 確認

▶ 能量源分類（62368-1 核心邏輯）
• 是否有做能量源分類章節（ES1/ES2/ES3）
• 缺此章節 = 高機率套用舊模板未更新 → WARN

▶ 關鍵量測項目（ee-test-report-review 檢查清單）
• 介電強度/耐壓：是否附實測漏電流數值；測試電壓與規格是否相符
• 接地阻抗：是否列出實測電阻值（通常 ≤0.1Ω，含 25-30A 測試電流）
• 漏電流：測試條件(單一/雙重故障、110%額定電壓)是否註明
• 爬電距離與電氣間隙：有無實測值對照；PCB 版本與送測樣品一致否
• 溫升測試：環境溫度、負載條件、高風險零件是否涵蓋
• 異常狀態測試：僅做部分項目 → WARN；缺元件故障模擬 → WARN
• 機構防護：UL94 認證文件(黃卡)是否附且未過期

▶ 常見 ODM 造假手法（ee-test-report-review 已知）
• 借用舊報告改標題頁
• 測試條件與出貨規格不符
• 只有 Summary 頁無 Raw Data
• 簽署與核准頁缺失

▶ 判定
• 若有上述 🔴 項 → FAIL
• 缺漏關鍵測項但給 PASS → CONTRADICTION''',

    'ENERGY_EFFICIENCY': '''【ENERGY_EFFICIENCY 能效測試】ENERGY STAR / EU ErP + ee-test-report-review

▶ 報告完整度（ee-test-report-review 要點）
• 是否只有計算結果無原始功率量測數據(raw power meter log) → WARN
• 功耗造假成本低，比安規/EMC 更需要原始數據佐證
• 規範版本是否為目的市場現行有效版本 → 需確認或 web search

▶ 核心量測項（ee-test-report-review 檢查清單）
• On Mode Power：亮度設定是否為出廠預設（取巧用低亮度測試 → WARN）；測試畫面內容是否符規範指定
• Sleep Mode Power：網路喚醒/藍芽是否維持出廠預設啟用狀態
• Off/Standby Power：是否涵蓋所有「準關機」狀態的說明
• 自動亮度控制/自動關閉：功能是否確實存在於出貨韌體（非僅測試版本）

▶ 常見 ODM 疏漏（ee-test-report-review 已知手法）
• 亮度設定取巧（低於出廠預設）→ WARN
• 測試畫面非規範指定內容 → WARN
• 韌體版本與出貨不一致 → WARN
• 只有 On Mode 無其他兩態 → 報告範疇不完整
• 規範版本過期（各版本限值可能放寬或收緊）→ WARN

▶ 判定
• 資料不完整或取巧跡象 → WARN
• 無法確認與出貨版本一致 → FAIL''',

    'DERATING': '''【DERATING 降額分析規則】

▶ 極重要：欄位對應自我驗證
• 用「實測值」÷「規格值」= 百分比 必須相等
• 若算不出來 → 你認錯欄位了，重新對應

▶ 降額標準
| 元件 | 降額要求 |
| 電阻(功率) | ≤ 額定×50% |
| 電容(電壓) | ≤ 額定×80% |
| 電感(電流) | ≤ 額定×80% |
| 二極體(電流) | ≤ 額定×75% |
| MOSFET(電壓) | ≤ 額定×80% |

▶ 多級門檻判定
• < Warning門檻 → PASS
• Warning ~ Waste間 → WARN
• ≥ Waste門檻 → FAIL

▶ 務必逐條檢查
1. 測試溫度是否為最惡劣工況
2. Max Load 定義是否涵蓋所有子系統''',

    'CRYSTAL': '''【CRYSTAL 晶振頻偏規則】

▶ 限值以各晶振 Datasheet 為準（通常 ±20 ppm 或 ±30 ppm）

▶ 判定
• |頻偏| > 規格限值 → FAIL
• |頻偏| > 80%限值 → WARN''',

    'PWR_TIMING': '''【PWR_TIMING 電源時序規則】

▶ 上電/下電
• 順序必符合 SoC/面板規格書
• 順序錯誤 → FAIL

▶ 各節點延遲
• 超出規格 → FAIL
• 裕量 < 10% → WARN''',

    'USB2_HUB': '''【USB2_HUB USB 2.0 Hub 信號完整性】

▶ Eye Diagram：Mask Hits = 0 → PASS；任何 Hits > 0 → FAIL
▶ 差分輸出電壓(HS)：400～600 mV → PASS
▶ 上升/下降時間：≤ 500 ps → PASS
▶ Margin < 10% → WARN''',

    'USB2_DEVICE': '''【USB2_DEVICE USB 2.0 Device 信號完整性】

▶ Eye Diagram：Mask Hits = 0 → PASS；任何 Hits > 0 → FAIL
▶ 差分輸出電壓(HS)：400～600 mV → PASS
▶ 上升/下降時間：≤ 500 ps → PASS''',

    'USB3_HOST': '''【USB3_HOST USB 3.2 Gen1/Gen2 信號完整性】

▶ Summary Failed = 0 → PASS；> 0 → FAIL
▶ 5G眼圖規格：Short/Far End DJ ≤ 86 ps；Mask Hits = 0
▶ 即使封面寫 PASS 但 Summary Failed > 0 → CONTRADICTION''',

    'RELIABILITY_CVTE': '''【RELIABILITY_CVTE 可靠性試驗】GB/T 2423

▶ 報告填寫判定
| 填寫 | 判定 |
| PASS | PASS |
| FAIL | FAIL |
| N/A + 說明 | 正常 |
| N/A 無說明 | WARN |
| 空白 | WARN |

▶ 矛盾檢測
• 封面 PASS 但逐項任一 FAIL → CONTRADICTION''',

    'DIFF_IMPEDANCE': '''【DIFF_IMPEDANCE 差分阻抗規則】

| 介面 | 規格 |
| HDMI 1.4/2.0 Thru | 85~115 Ω |
| HDMI 2.0 Term | 90~110 Ω |
| DisplayPort | 85~115 Ω |
| USB Type-C SuperSpeed | 72~118 Ω |

▶ 判定
• 超出規格 → FAIL
• 距邊界 < 3Ω → WARN''',

    'TEMP_RISE': '''【TEMP_RISE 溫升試驗】IEC 62368-1

▶ 報告完整度
• 測試時長、溫度穩態、工況說明是否齊全 → 確認

▶ 元件溫度限值
| 元件 | 限值 |
| 電感/變壓器(E級) | ≤ 130°C |
| 電解電容(105°C規格) | ≤ 105°C |
| 光耦(PC817類) | ≤ 100°C |

▶ 裕量計算
• 裕量% = (限值-實測值)/限值×100%
• 裕量% < 10% → WARN''',

    'ELEC': '''【ELEC 電氣性能規則】

▶ 六個子模塊檢查
• 屏時序測試：T1/T2/T3 是否在規格範圍；裕量<10% → WARN
• 開關機時序：各節點 > 100 ms → PASS，< 100 ms → FAIL
• 音頻測試：THD+N < 3% ✓；SNR ≥ 60dB ✓
• 聲學測試：噪聲 < 25dB(A) ✓；22~25dB → WARN
• 眼圖測試：所有通道 PASS → PASS；任一 FAIL → FAIL
• 時鐘數據：POWER ON > 0.7×VDD；POWER OFF < 0.3×VDD''',

    'VOLTAGE_RIPPLE': '''【VOLTAGE_RIPPLE 電壓紋波規則】

▶ 直流電壓準確度
• 在標稱值±5%內 → PASS
• 超出 → FAIL
• 裕量<1% → WARN

▶ 紋波(Ripple Vpp)
• 實測 > 限值 → FAIL
• 實測 > 80%限值 → WARN

▶ 裕量百分比
• 裕量% = (限值-實測值)/限值×100%
• 裕量% < 10% → WARN''',

    'OTA_WIRELESS': '''【OTA_WIRELESS OTA 無線性能】CTIA

| 頻段 | 參數 | 限值 |
| WiFi 2.4G/5G | TRP | ≥ 2 dBm |
| WiFi 2.4G/5G | TIS | ≤ -50 dBm |
| Bluetooth | TRP | ≥ -6 dBm |

▶ 「未測到」處理
• 數值超標 → FAIL
• 「未測到」無說明 → WARN
• 報告是否注明對應地區 → WARN''',

    'SOFTWARE_FUNCTIONAL': '''【SOFTWARE_FUNCTIONAL 軟體功能測試】

▶ 基本判定
| 結果 | 判定 |
| Pass/通過/OK | PASS |
| Fail/失敗/NG | FAIL |
| TBD/待確認/N/A(無說明) | WARN |

▶ 共因性缺陷聚合
• 同一故障在多通道重複 → 聚合說明，不逐條列為 N 個 FAIL

▶ 缺陷等級與 Pass/Fail 分離
• 先列整體比例
• 再獨立列缺陷等級(A級須優先)''',

    'VPC_RELIABILITY': '''【VPC_RELIABILITY VPC 環境可靠性試驗】

▶ 報告結構
• 「產品信息」：多 SKU 橫向呈現
• 逐項檢查 PASS/FAIL（ee-test-report-review Metadata 要點）

▶ 多 SKU 代表性
• 不同 CPU/主板/BIOS 配置是否都有獨立測試 → 檢查
• 僅測部分但宣稱全部 → WARN''',

    'VPC_STORAGE_STRESS': '''【VPC_STORAGE_STRESS VPC 硬碟壓力測試】

▶ 判定邏輯
• 逐 SKU、逐測項核對 PASS/FAIL
• 任一測項 FAIL 但總結 PASS → CONTRADICTION

▶ 儲存元件核心檢查
• 判定僅「開機時間」無資料完整性驗證 → WARN
• BurnInTest 是否列出具體項目；籠統「PASS」→ WARN''',

    'OTHER': '''【OTHER 自動識別模式】

▶ 從報告標題和內容自行判斷類型
▶ 套用對應標準審核規則
▶ 結果開頭說明：「識別為 XXX 測試報告，套用 XXX 審核規則。」'''
}

def extract_pdf_text(pdf_file) -> str:
    try:
        text = ""
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text += page.extract_text() + "\n"
        return text
    except Exception as e:
        logger.error(f"PDF 提取失敗: {str(e)}")
        return ""

def call_ai_api(system_prompt: str, user_prompt: str) -> str:
    """嘗試所有可用的 API"""
    apis = [
        ('Groq', GROQ_API_KEY, 'https://api.groq.com/openai/v1/chat/completions', 'mixtral-8x7b-32768'),
        ('OpenRouter', OPENROUTER_API_KEY, 'https://openrouter.ai/api/v1/chat/completions', 'openai/gpt-3.5-turbo'),
    ]
    
    for name, api_key, url, model in apis:
        if not api_key:
            continue
        try:
            headers = {
                'Authorization': f'Bearer {api_key}',
                'Content-Type': 'application/json',
                'HTTP-Referer': 'https://ifp-review-backend.onrender.com',
            }
            data = {
                'model': model,
                'messages': [
                    {'role': 'system', 'content': system_prompt},
                    {'role': 'user', 'content': user_prompt}
                ],
                'temperature': 0.7,
                'max_tokens': 2000
            }
            response = requests.post(url, json=data, headers=headers, timeout=60)
            if response.status_code == 200:
                result = response.json()
                if 'choices' in result and len(result['choices']) > 0:
                    logger.info(f"✅ {name} API 成功")
                    return result['choices'][0]['message']['content']
        except Exception as e:
            logger.warning(f"{name} API 錯誤: {str(e)}")
    
    logger.error("❌ 所有 API 都不可用")
    return "❌ API 不可用"

class ExcelExporter:
    def __init__(self, data: dict):
        self.data = data
        self.thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

    def split_content(self, text: str, max_width: int = 80) -> list:
        """將長文本按段落分割，保留換行符"""
        if not text:
            return ['']
        lines = text.split('\n')
        result = []
        for line in lines:
            if len(line) > max_width:
                # 按字符長度分割
                for i in range(0, len(line), max_width):
                    result.append(line[i:i+max_width])
            else:
                result.append(line)
        return result if result else ['']

    def export(self) -> tuple:
        wb = Workbook()
        ws = wb.active
        ws.title = "審核結果"
        
        # 設置列寬
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 90
        
        # 樣式定義
        title_font = Font(bold=True, size=16, color="FFFFFF")
        title_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        
        header_font = Font(bold=True, size=12, color="FFFFFF")
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        
        subheader_font = Font(bold=True, size=11, color="1F4E78")
        subheader_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        
        label_font = Font(bold=True, size=11, color="1F4E78")
        label_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
        
        content_font = Font(size=10)
        
        verdict = self.data.get('verdict', 'UNKNOWN')
        risk_level = self.data.get('risk_level', 'N/A')
        
        verdict_color = {
            "PASS": "70AD47",
            "FAIL": "FF0000",
            "WARN": "FFC000",
            "CONTRADICTION": "FF6600"
        }.get(verdict, "7F7F7F")
        
        row = 1
        
        # 【區塊一】標題
        title_cell = ws[f'A{row}']
        title_cell.value = 'IFP 測試報告審核結果'
        title_cell.font = title_font
        title_cell.fill = title_fill
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells(f'A{row}:B{row}')
        ws.row_dimensions[row].height = 28
        row += 1
        row += 1
        
        # 【區塊二】基本信息
        ws[f'A{row}'] = '基本信息'
        ws[f'A{row}'].font = header_font
        ws[f'A{row}'].fill = header_fill
        ws.merge_cells(f'A{row}:B{row}')
        ws.row_dimensions[row].height = 20
        row += 1
        
        info_items = [
            ('報告檔名', self.data.get('filename', 'N/A')),
            ('報告類型', self.data.get('report_type', 'N/A')),
            ('審核時間', datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
        ]
        
        for label, value in info_items:
            ws[f'A{row}'] = label
            ws[f'A{row}'].font = label_font
            ws[f'A{row}'].fill = label_fill
            ws[f'A{row}'].border = self.thin_border
            ws[f'A{row}'].alignment = Alignment(horizontal='left', vertical='center')
            
            ws[f'B{row}'] = value
            ws[f'B{row}'].font = content_font
            ws[f'B{row}'].border = self.thin_border
            ws[f'B{row}'].alignment = Alignment(wrap_text=True, vertical='center')
            ws.row_dimensions[row].height = 20
            row += 1
        
        row += 1
        
        # 【區塊三】最終判定
        ws[f'A{row}'] = '最終判定'
        ws[f'A{row}'].font = header_font
        ws[f'A{row}'].fill = header_fill
        ws.merge_cells(f'A{row}:B{row}')
        ws.row_dimensions[row].height = 20
        row += 1
        
        verdict_map = {
            'PASS': '✅ PASS - 通過',
            'FAIL': '❌ FAIL - 不通過',
            'WARN': '⚠️ WARN - 需澄清',
            'CONTRADICTION': '🔴 CONTRADICTION - 矛盾',
        }
        
        ws[f'A{row}'] = '判定'
        ws[f'A{row}'].font = label_font
        ws[f'A{row}'].fill = label_fill
        ws[f'A{row}'].border = self.thin_border
        ws[f'A{row}'].alignment = Alignment(horizontal='left', vertical='center')
        
        ws[f'B{row}'] = verdict_map.get(verdict, verdict)
        ws[f'B{row}'].font = Font(bold=True, size=12, color="FFFFFF")
        ws[f'B{row}'].fill = PatternFill(start_color=verdict_color, end_color=verdict_color, fill_type="solid")
        ws[f'B{row}'].border = self.thin_border
        ws[f'B{row}'].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[row].height = 22
        row += 1
        
        ws[f'A{row}'] = '風險燈號'
        ws[f'A{row}'].font = label_font
        ws[f'A{row}'].fill = label_fill
        ws[f'A{row}'].border = self.thin_border
        ws[f'A{row}'].alignment = Alignment(horizontal='left', vertical='center')
        
        ws[f'B{row}'] = risk_level
        ws[f'B{row}'].font = Font(bold=True, size=11)
        ws[f'B{row}'].border = self.thin_border
        ws[f'B{row}'].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[row].height = 20
        row += 2
        
        # 【區塊四】詳細內容 - 按段落分割顯示
        sections = [
            ('分析過程', 'analysis_process'),
            ('主要發現', 'main_findings'),
            ('詳細評論', 'detailed_comments'),
            ('ODM 追問清單', 'odm_questions'),
            ('風險等級判定', 'risk_summary'),
            ('重要提醒', 'important_note'),
        ]
        
        for section_title, section_key in sections:
            # 章節標題
            ws[f'A{row}'] = f'【{section_title}】'
            ws[f'A{row}'].font = subheader_font
            ws[f'A{row}'].fill = subheader_fill
            ws[f'A{row}'].border = self.thin_border
            ws.merge_cells(f'A{row}:B{row}')
            ws.row_dimensions[row].height = 20
            row += 1
            
            # 內容 - 按行分割
            content = self.data.get(section_key, '')
            if content:
                # 分割成多行，每行為一個 cell
                lines = self.split_content(content, max_width=85)
                for idx, line in enumerate(lines):
                    ws[f'A{row}'] = '' if idx > 0 else ''  # 第一行空，後面行也空
                    ws[f'B{row}'] = line
                    ws[f'B{row}'].font = content_font
                    ws[f'B{row}'].border = self.thin_border
                    ws[f'B{row}'].alignment = Alignment(wrap_text=False, vertical='top')
                    ws.row_dimensions[row].height = 18
                    row += 1
            else:
                ws[f'B{row}'] = '(無)'
                ws[f'B{row}'].font = Font(size=10, italic=True, color="999999")
                ws[f'B{row}'].border = self.thin_border
                ws.row_dimensions[row].height = 18
                row += 1
            
            row += 1
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        filename = f"{self.data.get('filename', 'report').replace('.pdf', '')}_審核結果.xlsx"
        return output.getvalue(), filename

@app.route('/', methods=['GET'])
def index():
    try:
        with open('index.html', 'r', encoding='utf-8') as f:
            return f.read()
    except:
        return "前端未找到", 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'version': '9.4-Merged'})

@app.route('/api/generate-prompt', methods=['POST'])
def gen_prompt():
    data = request.get_json()
    report_type = data.get('type', 'OTHER')
    stage = data.get('stage', 'DVT')
    
    type_rules = TYPE_SPECIFIC_RULES.get(report_type, TYPE_SPECIFIC_RULES['OTHER'])
    
    prompt = f'''{UNIVERSAL_FRAMEWORK}

【測試階段】{stage}
【報告類型】{report_type}

【{report_type} 專用審核規則】

{type_rules}

---

【請粘貼報告文本】
'''
    
    return jsonify({'status': 'success', 'prompt': prompt})

@app.route('/api/review', methods=['POST'])
def review():
    """自動模式 - 上傳 PDF 自動分析"""
    try:
        if 'file' not in request.files:
            return jsonify({'status': 'error', 'message': '缺少文件'}), 400
        
        file = request.files['file']
        if not file or not file.filename.endswith('.pdf'):
            return jsonify({'status': 'error', 'message': '只支持 PDF 文件'}), 400
        
        report_type = request.form.get('type', 'AUTO')
        stage = request.form.get('stage', 'DVT')
        
        pdf_text = extract_pdf_text(file.stream)
        if not pdf_text:
            return jsonify({'status': 'error', 'message': 'PDF 文本提取失敗'}), 400
        
        type_rules = TYPE_SPECIFIC_RULES.get(report_type, TYPE_SPECIFIC_RULES['OTHER'])
        
        system_prompt = f'''{UNIVERSAL_FRAMEWORK}

【{report_type} 專用規則】
{type_rules}

請簡潔分析報告。'''

        user_prompt = f"""測試階段：{stage}
報告類型：{report_type}

請分析報告文本（前 8000 字）：

{pdf_text[:8000]}

最後明確說明最終判定：PASS 或 FAIL 或 WARN 或 CONTRADICTION"""

        result = call_ai_api(system_prompt, user_prompt)
        
        verdict = 'PASS'
        upper = result.upper()
        if 'CONTRADICTION' in upper:
            verdict = 'CONTRADICTION'
        elif 'FAIL' in upper:
            verdict = 'FAIL'
        elif 'WARN' in upper:
            verdict = 'WARN'
        
        return jsonify({
            'status': 'success',
            'data': {
                'filename': file.filename,
                'report_type': report_type,
                'stage': stage,
                'verdict': verdict,
                'analysis_process': result,
                'main_findings': result,
                'detailed_comments': result,
                'odm_questions': result,
                'risk_summary': result,
                'risk_level': {'PASS': '🟢', 'FAIL': '🔴', 'WARN': '🟡', 'CONTRADICTION': '🔴'}.get(verdict, '🟢'),
                'timestamp': datetime.now().isoformat()
            }
        })
    
    except Exception as e:
        logger.error(f"分析錯誤: {str(e)}")
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/export/excel', methods=['POST'])
def export():
    data = request.get_json() or {}
    
    safe_data = {
        'filename': data.get('filename', 'report.pdf'),
        'report_type': data.get('report_type', 'OTHER'),
        'verdict': data.get('verdict', 'PASS'),
        'risk_level': data.get('risk_level', '🟢'),
        'analysis_process': data.get('analysis_process', ''),
        'main_findings': data.get('main_findings', ''),
        'detailed_comments': data.get('detailed_comments', ''),
        'odm_questions': data.get('odm_questions', ''),
        'risk_summary': data.get('risk_summary', ''),
        'important_note': data.get('important_note', '本審核為 AI 輔助之文件初篩，不能取代具資格 EE 工程師的技術判定。'),
    }
    
    exporter = ExcelExporter(safe_data)
    content, filename = exporter.export()
    return send_file(BytesIO(content), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=filename)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)), debug=False)

