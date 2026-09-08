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

【重要】你的輸出不是文本，而是 JSON 代碼塊。用戶將復制這個 JSON 到工具中生成 Excel。
因此，用戶在 Claude.ai 看不到完整分析過程，只能看到 JSON 代碼。

【JSON 輸出格式】
```json
{
  "filename": "[報告檔名]",
  "report_type": "[報告類型]",
  "verdict": "PASS|FAIL|WARN|CONTRADICTION",
  "risk_level": "🔴 高風險|🟡 需澄清|🟢 通過",
  "analysis_process": "【分析過程】詳細步驟...",
  "main_findings": "【主要發現】核心問題...",
  "detailed_comments": "【詳細評論】完整評論...",
  "odm_questions": "【回覆 ODM 追問清單】\n1. ...\n2. ...",
  "risk_summary": "【風險等級判定】🟡 條件 PASS...",
  "important_note": "【重要提醒】本審核為 AI 輔助..."
}
```

【說明】
- 只在 JSON 的 value 中放完整分析內容
- 用戶看不到分析過程（只看到 JSON 代碼塊）
- 用戶必須復制這個 JSON 到工具才能查看完整內容並生成 Excel
'''

# ===== 20種報告類型專用規則 - 精簡版 =====
TYPE_SPECIFIC_RULES = {
    'DERATING': '''【DERATING 降額分析規則】
▶ 極重要：用「實測值」÷「規格值」必須相等
▶ 降額標準：電阻≤50%、電容≤80%、電感≤80%、二極體≤75%、MOSFET≤80%
▶ 判定：< Warning → PASS；Warning~Waste → WARN；≥ Waste → FAIL''',

    'SAFETY_IEC62368': '''【SAFETY_IEC62368 IEC 62368-1 安規測試】
▶ 是否引用舊標準(60950-1)？ → WARN
▶ 標明標準版次、額定規格變更複測、PS/ES分類？ → WARN
▶ 接觸電流(次級)<1mA、(金屬)<3.5mA；USB<8A；耐壓3000VAC''',

    'EMI': '''【EMI 電磁騷擾規則】EN 55032 Class B
▶ 限值與Margin：30～230 MHz: QP ≤ 40 dBµV/m；Margin < 6 dB → WARN
▶ 判定：Margin ≤ 0 dB → FAIL；封面PASS但任何頻點超標 → CONTRADICTION''',

    'OTHER': '''【OTHER 自動識別模式】
▶ 從報告標題和測試內容自行判斷類型，套用對應標準審核規則'''
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

@app.route('/api/version', methods=['GET'])
def get_version():
    return jsonify({
        'version': '9.1-Final',
        'mode': '手動模式：JSON 代碼 → 工具生成 Excel',
        'framework': 'Merged RD 7-Step + EE-Test-Report-Review'
    })

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=FLASK_PORT, debug=DEBUG_MODE)

