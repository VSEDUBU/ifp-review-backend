#!/usr/bin/env python3
"""
IFP 測試報告自動審核工具 v9.1 Final
- 手動模式：Prompt 輸出 JSON 代碼，用戶復制粘貼到工具生成 Excel
- 自動模式：支持 4 個 API
"""

import os
import json
import logging
from datetime import datetime
from io import BytesIO

import requests
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from dotenv import load_dotenv
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

load_dotenv()

GROQ_API_KEY = os.getenv('GROQ_API_KEY')
OPENROUTER_API_KEY = os.getenv('OPENROUTER_API_KEY')
CEREBRAS_API_KEY = os.getenv('CEREBRАС_API_KEY') or os.getenv('CEREBRAS_API_KEY')
GEMINI_API_KEY = os.getenv('GEMINI_API_KEY')

FLASK_PORT = int(os.getenv('FLASK_PORT', 5000))
DEBUG_MODE = os.getenv('DEBUG_MODE', 'False').lower() == 'true'

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

# ===== 通用審核框架 - 融合版 =====
UNIVERSAL_FRAMEWORK = '''【IFP 測試報告審核框架 - 融合版】

你是資深硬體工程師，負責審核測試報告。

【RD 級審核 - 7 步】

1️⃣ 理解報告
   • 識別報告類型：電氣安全(Safety/IEC 62368-1) / EMC/EMI(CISPR 32/35) / 能效(Energy Star/DOE/EU ErP)
   • 確認適用標準、版本、補充規範

2️⃣ 驗證條件 + Metadata 檢查（ee-test-report-review）
   • 型號/PN/版本：與送審 BOM 或送測樣品是否一致？
   • 報告號/日期/頁數：是否完整、有無跳頁或格式突變？
   • 測試實驗室：是否具備該項目認證（ISO 17025/CNAS/A2LA/CB Scheme）？認證是否在有效期內？
   • 報告簽署：簽署人、審核人、簽名/章是否完整？
   • 樣品照片：銘牌 label photo 是否與型號相符？

3️⃣ 檢查數據 - 規格、實測、判定邏輯是否一致？
4️⃣ 覆蓋完整性 - 應測項目是否都測了？
5️⃣ 矛盾檢測 - 結論與數據一致性？
6️⃣ 風險評估 + 風險燈號
   • 🔴 高風險（阻擋簽核）
   • 🟡 需澄清
   • 🟢 通過

7️⃣ 最終判定 + ODM 追問清單

【關鍵指示 - 必須遵守】
你的輸出 **只能是下面的 JSON 代碼塊**，別無他物。
- ❌ 不要創建文件
- ❌ 不要下載
- ❌ 不要詢問用戶
- ❌ 代碼塊前面不要有任何文字
- ❌ 代碼塊後面不要有任何文字
- ✅ 只輸出純文本 JSON

【直接生成 Excel - Claude.ai Artifact】
不要輸出任何分析文本，直接在 Artifact 創建可下載的 Excel 文件。

【Claude 務必複製以下完整代碼到 Artifact】

```html
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.min.js"></script>
    <style>
        body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
        .container { max-width: 900px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; }
        h1 { color: white; background: #1F4E78; padding: 20px; border-radius: 6px; }
        button { padding: 12px 24px; background: #4472C4; color: white; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; width: 100%; margin-top: 20px; }
        button:hover { background: #2F5496; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📋 IFP 測試報告審核結果</h1>
        <p id="info"></p>
        <button onclick="downloadExcel()">📥 下載 Excel 審核報告</button>
    </div>
    <script>
        const data = { filename: "[報告檔名]", report_type: "[報告類型]", verdict: "[PASS]", risk_level: "[🟢]", analysis_process: "[分析過程]", main_findings: "[主要發現]", detailed_comments: "[詳細評論]", odm_questions: "[ODM追問]", risk_summary: "[風險判定]", important_note: "[重要提醒]" };
        document.getElementById('info').innerHTML = `<strong>報告：</strong> ${data.filename}<br><strong>類型：</strong> ${data.report_type}<br><strong>判定：</strong> ${data.verdict}`;
        function downloadExcel() {
            try {
                const wb = XLSX.utils.book_new();
                const ws = XLSX.utils.aoa_to_sheet([['IFP 測試報告審核結果'], [], ['報告檔名', data.filename], ['報告類型', data.report_type], ['最終判定', data.verdict], ['風險燈號', data.risk_level], [], ['分析過程'], [data.analysis_process], [], ['主要發現'], [data.main_findings], [], ['詳細評論'], [data.detailed_comments], [], ['ODM追問'], [data.odm_questions], [], ['風險判定'], [data.risk_summary], [], ['重要提醒'], [data.important_note]]);
                ws['!cols'] = [{wch: 25}, {wch: 100}];
                XLSX.utils.book_append_sheet(wb, ws, "審核結果");
                
                // 使用 Blob 方式下载
                const wbout = XLSX.write(wb, {bookType: 'xlsx', type: 'array'});
                const blob = new Blob([wbout], {type: 'application/octet-stream'});
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = data.filename.replace('.pdf', '') + '_審核結果.xlsx';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            } catch (error) {
                alert('下載失敗：' + error.message);
            }
        }
    </script>
</body>
</html>
```

【Claude 的工作】
1. 複製上面的 HTML 到 Artifact
2. 分析後填入 const data 中的各欄位
3. 用戶點擊【📥 下載 Excel 審核報告】直接下載
'''

# ===== 20種報告類型專用規則 - 完整版 =====
TYPE_SPECIFIC_RULES = {
    'EMI': '''【EMI 電磁騷擾規則】EN 55032 Class B

▶ 報告完整度
• 若摘要判FAIL，內頁必有對應頻點詳細數據表
• 測試工況：是否涵蓋整機最惡劣（最高解析度+滿載周邊）
• Class A/B誤用：IFP應用Class B（較嚴格）

▶ 限值與Margin
• 30～230 MHz：QP ≤ 40 dBµV/m
• 230～1000 MHz：QP ≤ 47 dBµV/m
• 1～3 GHz：PK ≤ 70 / AV ≤ 50 dBµV/m
• 3～6 GHz：PK ≤ 74 / AV ≤ 54 dBµV/m
• 傳導騷擾(CE)：150k～500kHz ≤56 / 0.5M～5MHz ≤66 / 5M～30MHz ≤60 dBµV

▶ 判定
• Margin < 6 dB → WARN
• Margin ≤ 0 dB → FAIL
• 封面PASS但任何頻點超標 → CONTRADICTION''',

    'EMS': '''【EMS 電磁抗擾度規則】EN 55035 / IEC 61000-4

▶ 報告完整度
• 必須有逐項測試結果表（ESD/EFT/Surge/CS/Dips各自數據）
• 僅有等級定義無具體測試結果 → WARN

▶ 等級定義
• A：完全正常，無任何降級
• B：輕微降級（如閃屏），自動恢復
• C：需手動恢復（重啟）
• D：無法恢復（硬件損傷）

▶ 各測試項最低要求
| 項目 | 最低 | PASS條件 |
| ESD | B | A或B |
| RS輻射 | A | 僅A |
| EFT | B | A或B |
| Surge | B | A或B |
| CS傳導 | A | 僅A |
| Dips | C | A/B/C |

▶ 矛盾檢測
• 封面PASS但內頁任一項C或D → CONTRADICTION''',

    'ELEC': '''【ELEC 電氣性能規則】6個子模塊

▶ 子模塊一：屏時序測試(Panel Timing)
• T1(Vcc上升)：0.5~10 ms
• T2(Vcc到CDR)：≥50 ms
• T3(信號到Vcc關機)：≥0 ms
• 超出範圍 → FAIL，裕量<10% → WARN

▶ 子模塊二：開關機時序
• 各節點 > 100 ms → PASS，< 100 ms → FAIL

▶ 子模塊三：音頻測試
• THD+N < 3% ✓，SNR ≥ 60dB ✓

▶ 子模塊四：聲學測試
• 噪聲 < 25dB(A) ✓；22~25dB → WARN

▶ 子模塊五：眼圖測試
• 所有通道PASS → PASS；任一FAIL → FAIL

▶ 子模塊六：時鐘數據
• POWER ON > 0.7×VDD ✓
• POWER OFF < 0.3×VDD ✓''',

    'VOLTAGE_RIPPLE': '''【VOLTAGE_RIPPLE 電壓紋波規則】

▶ 直流電壓準確度
• 在標稱值±5%內 → PASS
• 超出 → FAIL
• 裕量<1% → WARN

▶ 紋波(Ripple Vpp)限值
• 實測 > 限值 → FAIL
• 實測 > 80%限值 → WARN

▶ 裕量百分比(必須逐節點計算)
• 裕量% = (限值-實測值)/限值×100%
• 裕量% < 10% → WARN''',

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
1. 測試溫度是否為最惡劣工況？
2. Max Load定義是否涵蓋所有子系統？''',

    'CRYSTAL': '''【CRYSTAL 晶振頻偏規則】

▶ 限值以各晶振Datasheet為準
• 通常 ±20 ppm 或 ±30 ppm

▶ 判定
• |頻偏| > 規格限值 → FAIL
• |頻偏| > 80%限值 → WARN''',

    'PWR_TIMING': '''【PWR_TIMING 電源時序規則】

▶ 上電/下電
• 順序必符合SoC/面板規格書
• 順序錯誤 → FAIL

▶ 各節點延遲
• 超出規格 → FAIL
• 裕量 < 10% → WARN''',

    'USB2_HUB': '''【USB2_HUB USB 2.0 Hub 信號完整性】

▶ Eye Diagram
• Mask Hits = 0 → PASS
• 任何 Hits > 0 → FAIL

▶ 差分輸出電壓(HS)
• 400 mV ≤ VDiff ≤ 600 mV → PASS

▶ 上升/下降時間
• ≤ 500 ps → PASS

▶ Margin < 10% → WARN''',

    'USB2_DEVICE': '''【USB2_DEVICE USB 2.0 Device 信號完整性】

▶ Eye Diagram
• Mask Hits = 0 → PASS
• 任何 Hits > 0 → FAIL

▶ 差分輸出電壓(HS)
• 400 mV ≤ VDiff ≤ 600 mV → PASS

▶ 上升/下降時間
• ≤ 500 ps → PASS''',

    'USB3_HOST': '''【USB3_HOST USB 3.2 Gen1/Gen2 信號完整性】

▶ Summary Failed
• = 0 → PASS
• > 0 → FAIL（即使封面寫Pass → CONTRADICTION）

▶ LFPS時序規格
• LFPS差分電壓：800~1200 mV
• Rise/Fall Time：≤ 4.0 ns

▶ 5G眼圖規格
• Short/Far End DJ：≤ 86 ps
• Mask Hits：= 0''',

    'RELIABILITY_CVTE': '''【RELIABILITY_CVTE 可靠性試驗】GB/T 2423或企業標準

▶ 報告填寫判定
| 填寫 | 判定 |
| PASS | PASS |
| FAIL | FAIL |
| N/A + 說明 | NA(正常) |
| N/A 無說明 | WARN |
| 空白 | WARN |

▶ 矛盾檢測(務必逐項核對)
• 封面PASS但逐項清單任一子項FAIL → CONTRADICTION
• 封面FAIL但逐項清單全PASS → CONTRADICTION''',

    'DIFF_IMPEDANCE': '''【DIFF_IMPEDANCE 差分阻抗規則】

▶ 規格範圍
| 介面 | 規格 |
| HDMI 1.4/2.0 Thru | 85~115 Ω |
| HDMI 2.0 Term | 90~110 Ω |
| DisplayPort | 85~115 Ω |
| USB Type-C SuperSpeed | 72~118 Ω |

▶ 判定
• 超出規格 → FAIL
• 距邊界 < 3Ω → WARN''',

    'TEMP_RISE': '''【TEMP_RISE 溫升試驗】IEC 62368-1 Annex M

▶ 報告完整度
• 是否記錄測試時長？未記錄 → WARN
• 是否達到溫度穩態？ → 確認
• 是否記錄測試工況？ → 確認

▶ 元件溫度限值
| 元件 | 限值 |
| 電感/變壓器(E級) | ≤ 130°C |
| 電解電容(105°C規格) | ≤ 105°C |
| 光耦(PC817類) | ≤ 100°C |

▶ 裕量計算方式
• 裕量% = (限值-實測值)/限值×100%
• 裕量% < 10% → WARN''',

    'SAFETY_IEC62368': '''【SAFETY_IEC62368 IEC 62368-1 安規測試】

▶ 報告完整度 (Metadata重點)
• 是否引用已被62368-1取代的舊標準？ → WARN
• 是否標明標準版次？ → WARN
• 額定規格(Rating)是否與BOM一致？變更後是否複測？ → 檢查
• PS/ES等分類等級是否標示？ → WARN
• 樣品照片(銘牌)是否清晰？ → 確認

▶ 測試項目與限值
| 項目 | 規格 |
| 接觸電流(次級) | < 1 mA |
| 接觸電流(金屬) | < 3.5 mA |
| USB埠電流 | < 8 A |
| 電氣強度 | 3000 VAC |
| 絕緣電阻 | ≥ 250 MΩ |
| 接地連續性 | ≤ 0.1 Ω |''',

    'ENERGY_EFFICIENCY': '''【ENERGY_EFFICIENCY 能效測試】EU 2019/2021 + EU 2019/2013

▶ EU 2019/2021(待機/關機功耗，強制)
| 模式 | 限值 |
| Standby mode | < 0.5 W |
| Off mode | < 0.3 W |
| Network standby | < 2.0 W |

▶ Metadata檢查
• 測試亮度設定是否 = 出廠預設？ → WARN
• 測試畫面是否為規範指定？ → WARN
• Peak Luminance Ratio若 < 65% 但標"參考用" → CONTRADICTION

▶ EU 2019/2013(能效標籤，必須檢查)
• 無論報告標記都必須列出等級
• F或G級 → WARN
• D或E級 → WARN''',

    'OTA_WIRELESS': '''【OTA_WIRELESS OTA 無線性能】CTIA

▶ 規格門檻
| 頻段 | 參數 | 限值 |
| WiFi 2.4G/5G | TRP | ≥ 2 dBm |
| WiFi 2.4G/5G | TIS | ≤ -50 dBm |
| Bluetooth | TRP | ≥ -6 dBm |
| Bluetooth | TIS | ≤ -50 dBm |

▶ "未測到"處理
• 數值超標 → FAIL
• "未測到" → WARN(可能測試治具異常)

▶ 法規符合性
• 報告是否注明對應地區？ → WARN''',

    'SOFTWARE_FUNCTIONAL': '''【SOFTWARE_FUNCTIONAL 軟體功能測試】

▶ 基本判定
| 結果 | 判定 |
| Pass/通過/OK | PASS |
| Fail/失敗/NG | FAIL |
| TBD/待確認/N/A(無說明) | WARN |
| N/A(明確說明) | NA |

▶ 共因性缺陷聚合
• 同一故障模式在多個通道重複 → 疑似共因
• 不要逐條列成N個FAIL，改為聚合說明

▶ 缺陷等級與Pass/Fail分離
• 先列整體Pass/Fail比例
• 再獨立列缺陷等級(A級須優先)''',

    'VPC_RELIABILITY': '''【VPC_RELIABILITY VPC 環境可靠性試驗】GB/T 2423或企業標準

▶ 報告結構
• "產品信息"工作表：多SKU橫向呈現
• 各環測子表：試驗配置、標準、溫濕條件、檢測項目

▶ 判定邏輯
• 逐一核對每一列PASS/FAIL
• 矛盾檢測同RELIABILITY_CVTE

▶ 多SKU代表性
• 不同CPU/主板/BIOS配置是否都有獨立測試？
• 僅測部分宣稱全部 → WARN''',

    'VPC_STORAGE_STRESS': '''【VPC_STORAGE_STRESS VPC 硬碟壓力測試】

▶ 報告結構
• "測試大綱及總結"：統計與缺陷等級
• "硬碟測試list"：SKU縱向，測項橫向

▶ 判定邏輯
• 逐SKU、逐測項核對PASS/FAIL
• 任一測項FAIL但總結PASS → CONTRADICTION

▶ 儲存元件核心檢查
1. 判定僅"開機時間"無資料完整性驗證 → WARN
2. BurnInTest是否列出具體測試項？籠統"PASS" → WARN''',

    'OTHER': '''【OTHER 自動識別模式】

▶ 從報告標題和測試內容自行判斷類型
▶ 套用對應標準審核規則
▶ 結果開頭說明："識別為 XXX 測試報告，套用 XXX 審核規則。"'''
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
                    logger.info(f"✅ 使用 {name} API")
                    return result['choices'][0]['message']['content']
        except Exception as e:
            logger.warning(f"{name} API 失敗: {str(e)}")
            continue
    
    return "❌ 所有 API 都不可用"

class ExcelExporter:
    def __init__(self, data: dict):
        self.data = data

    def export(self) -> tuple:
        wb = Workbook()
        ws = wb.active
        ws.title = "審核結果"
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 70
        
        title_font = Font(bold=True, size=16, color="1F4E78")
        section_font = Font(bold=True, size=12, color="FFFFFF")
        section_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        
        row = 1
        title_cell = ws[f'A{row}']
        title_cell.value = 'IFP 測試報告審核結果'
        title_cell.font = title_font
        ws.merge_cells(f'A{row}:B{row}')
        row += 2
        
        # 審核信息
        section_cell = ws[f'A{row}']
        section_cell.value = "📋 審核信息"
        section_cell.font = section_font
        section_cell.fill = section_fill
        ws.merge_cells(f'A{row}:B{row}')
        row += 1
        
        info_data = [
            ('報告檔名', self.data.get('filename', 'N/A')),
            ('報告類型', self.data.get('report_type', 'N/A')),
            ('審核時間', datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
            ('風險燈號', self.data.get('risk_level', 'N/A')),
        ]
        
        for label, value in info_data:
            ws[f'A{row}'] = label
            ws[f'A{row}'].font = Font(bold=True, size=11)
            ws[f'A{row}'].fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
            ws[f'B{row}'] = value
            row += 1
        
        row += 1
        
        # 最終判定
        verdict = self.data.get('verdict', 'UNKNOWN')
        verdict_color = {"PASS": "70AD47", "FAIL": "FF0000", "WARN": "FFC000", "CONTRADICTION": "FF6600"}.get(verdict, "7F7F7F")
        
        result_section = ws[f'A{row}']
        result_section.value = "📊 最終判定"
        result_section.font = section_font
        result_section.fill = section_fill
        ws.merge_cells(f'A{row}:B{row}')
        row += 1
        
        verdict_map = {
            'PASS': '✅ PASS',
            'FAIL': '❌ FAIL',
            'WARN': '⚠️ WARN',
            'CONTRADICTION': '🔴 CONTRADICTION',
        }
        
        ws[f'B{row}'] = verdict_map.get(verdict, verdict)
        ws[f'B{row}'].font = Font(bold=True, size=12, color="FFFFFF")
        ws[f'B{row}'].fill = PatternFill(start_color=verdict_color, end_color=verdict_color, fill_type="solid")
        row += 2
        
        # 分析內容
        sections = [
            ('分析過程', 'analysis_process'),
            ('主要發現', 'main_findings'),
            ('詳細評論', 'detailed_comments'),
            ('ODM 追問清單', 'odm_questions'),
            ('風險等級判定', 'risk_summary'),
            ('重要提醒', 'important_note'),
        ]
        
        for section_title, section_key in sections:
            section_cell = ws[f'A{row}']
            section_cell.value = f"📝 {section_title}"
            section_cell.font = section_font
            section_cell.fill = section_fill
            ws.merge_cells(f'A{row}:B{row}')
            row += 1
            
            content = self.data.get(section_key, '')
            ws[f'A{row}'] = content
            ws[f'A{row}'].alignment = Alignment(wrap_text=True, vertical='top')
            ws.merge_cells(f'A{row}:B{row}')
            ws.row_dimensions[row].height = 200
            row += 2
        
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
    except FileNotFoundError:
        return "前端文件未找到", 500

@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({
        'status': 'healthy',
        'version': '9.1-Final',
        'groq_configured': bool(GROQ_API_KEY),
        'openrouter_configured': bool(OPENROUTER_API_KEY),
        'timestamp': datetime.now().isoformat()
    })

@app.route('/api/generate-prompt', methods=['POST'])
def generate_prompt():
    try:
        data = request.get_json()
        report_type = data.get('type', 'OTHER')
        stage = data.get('stage', 'DVT')
        
        type_rules = TYPE_SPECIFIC_RULES.get(report_type, TYPE_SPECIFIC_RULES['OTHER'])
        
        prompt = f'''{UNIVERSAL_FRAMEWORK}

【測試階段】{stage}
【報告類型】{report_type}

【報告類型專用審核規則】

{type_rules}

---

【請在下方粘貼報告文本】
'''
        
        return jsonify({'status': 'success', 'prompt': prompt})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/export/excel', methods=['POST'])
def export_excel():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'status': 'error', 'message': '缺少數據'}), 400
        exporter = ExcelExporter(data)
        content, filename = exporter.export()
        return send_file(
            BytesIO(content),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

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

【報告類型專用規則】
{type_rules}

請簡潔分析報告。'''

        user_prompt = f"""測試階段：{stage}
報告類型：{report_type}

請分析報告文本（前 8000 字）：

{pdf_text[:8000]}

最後說明：最終判定：PASS 或 FAIL 或 WARN 或 CONTRADICTION"""

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
                'result': result,
                'timestamp': datetime.now().isoformat()
            }
        })
    
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/version', methods=['GET'])
def get_version():
    return jsonify({
        'version': '9.1-Final',
        'mode': '手動模式：JSON 代碼 → 工具生成 Excel',
        'framework': 'Merged RD 7-Step + EE-Test-Report-Review'
    })

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=FLASK_PORT, debug=DEBUG_MODE)

