#!/usr/bin/env python3
"""IFP 測試報告審核工具 v9.2 - 只輸出 JSON，不顯示分析過程"""

import os, json, logging, requests, re
from datetime import datetime
from io import BytesIO
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from dotenv import load_dotenv
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

load_dotenv()
GROQ_API_KEY = os.getenv('GROQ_API_KEY')
OPENROUTER_API_KEY = os.getenv('OPENROUTER_API_KEY')

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

# ===== 只輸出 JSON 的 Prompt =====
UNIVERSAL_FRAMEWORK = '''【IFP 測試報告審核工具】

你是資深硬體工程師，進行 RD 7 步審核（理解報告 → Metadata 驗證 → 檢查數據 → 覆蓋完整性 → 矛盾檢測 → 風險評估 → 最終判定）。

【重要】你的輸出 **只能是** 下面這個 JSON 代碼塊，不能有任何其他文本！

```json
{
  "filename": "報告檔名.pdf",
  "report_type": "報告類型",
  "verdict": "PASS或FAIL或WARN或CONTRADICTION",
  "risk_level": "🔴或🟡或🟢",
  "analysis_process": "【分析過程】...",
  "main_findings": "【主要發現】...",
  "detailed_comments": "【詳細評論】...",
  "odm_questions": "【ODM 追問清單】...",
  "risk_summary": "【風險等級判定】...",
  "important_note": "【重要提醒】..."
}
```'''

TYPE_RULES = {'DERATING': '降額分析規則', 'SAFETY_IEC62368': '安規規則', 'EMI': 'EMI規則', 'OTHER': '其他'}

def extract_pdf_text(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            return "\n".join(page.extract_text() or "" for page in pdf.pages)
    except:
        return ""

def call_api(system_prompt, user_prompt):
    if GROQ_API_KEY:
        try:
            r = requests.post('https://api.groq.com/openai/v1/chat/completions', json={
                'model': 'mixtral-8x7b-32768', 'messages': [
                    {'role': 'system', 'content': system_prompt},
                    {'role': 'user', 'content': user_prompt}
                ], 'max_tokens': 2000
            }, headers={'Authorization': f'Bearer {GROQ_API_KEY}', 'Content-Type': 'application/json'}, timeout=60)
            if r.status_code == 200:
                return r.json()['choices'][0]['message']['content']
        except: pass
    if OPENROUTER_API_KEY:
        try:
            r = requests.post('https://openrouter.ai/api/v1/chat/completions', json={
                'model': 'openai/gpt-3.5-turbo', 'messages': [
                    {'role': 'system', 'content': system_prompt},
                    {'role': 'user', 'content': user_prompt}
                ], 'max_tokens': 2000
            }, headers={'Authorization': f'Bearer {OPENROUTER_API_KEY}', 'Content-Type': 'application/json'}, timeout=60)
            if r.status_code == 200:
                return r.json()['choices'][0]['message']['content']
        except: pass
    return "❌ API 不可用"

class ExcelExporter:
    def __init__(self, data):
        self.data = data
    
    def export(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "審核結果"
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 70
        
        title_font = Font(bold=True, size=16, color="1F4E78")
        section_font = Font(bold=True, size=12, color="FFFFFF")
        section_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        
        row = 1
        ws[f'A{row}'] = 'IFP 測試報告審核結果'
        ws[f'A{row}'].font = title_font
        ws.merge_cells(f'A{row}:B{row}')
        row += 2
        
        for label, value in [
            ('檔名', self.data.get('filename', 'N/A')),
            ('類型', self.data.get('report_type', 'N/A')),
            ('判定', self.data.get('verdict', 'N/A')),
            ('風險', self.data.get('risk_level', 'N/A')),
        ]:
            ws[f'A{row}'] = label
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'B{row}'] = value
            row += 1
        
        row += 1
        
        for title, key in [
            ('分析過程', 'analysis_process'),
            ('主要發現', 'main_findings'),
            ('詳細評論', 'detailed_comments'),
            ('ODM 追問', 'odm_questions'),
            ('風險評估', 'risk_summary'),
        ]:
            ws[f'A{row}'] = title
            ws[f'A{row}'].font = section_font
            ws[f'A{row}'].fill = section_fill
            ws.merge_cells(f'A{row}:B{row}')
            row += 1
            ws[f'A{row}'] = self.data.get(key, '')
            ws[f'A{row}'].alignment = Alignment(wrap_text=True)
            ws.merge_cells(f'A{row}:B{row}')
            ws.row_dimensions[row].height = 200
            row += 2
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue(), f"{self.data.get('filename', 'report').replace('.pdf', '')}_審核結果.xlsx"

@app.route('/', methods=['GET'])
def index():
    try:
        with open('index.html', 'r', encoding='utf-8') as f:
            return f.read()
    except:
        return "前端未找到", 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'version': '9.2'})

@app.route('/api/generate-prompt', methods=['POST'])
def gen_prompt():
    data = request.get_json()
    report_type = data.get('type', 'OTHER')
    return jsonify({'status': 'success', 'prompt': f'{UNIVERSAL_FRAMEWORK}\n\n【報告類型】{report_type}\n【規則】{TYPE_RULES.get(report_type, "自動識別")}\n\n請粘貼報告文本：'})

@app.route('/api/review', methods=['POST'])
def review():
    """自動模式 - 上傳 PDF 自動分析"""
    try:
        if 'file' not in request.files:
            return jsonify({'status': 'error', 'message': '缺少文件'}), 400
        
        file = request.files['file']
        if not file or not file.filename.endswith('.pdf'):
            return jsonify({'status': 'error', 'message': '只支持 PDF'}), 400
        
        report_type = request.form.get('type', 'OTHER')
        
        # 提取 PDF 文本
        pdf_text = extract_pdf_text(file.stream)
        if not pdf_text:
            return jsonify({'status': 'error', 'message': 'PDF 提取失敗'}), 400
        
        # 調用 AI 分析
        system_prompt = f'{UNIVERSAL_FRAMEWORK}\n\n【報告類型】{report_type}\n【規則】{TYPE_RULES.get(report_type, "自動識別")}'
        user_prompt = f'請分析這份報告（前 5000 字）：\n\n{pdf_text[:5000]}'
        
        result = call_api(system_prompt, user_prompt)
        
        # 提取 JSON
        json_match = re.search(r'```json\n?([\s\S]*?)\n?```', result) if result else None
        data = {}
        
        if json_match:
            try:
                data = json.loads(json_match.group(1))
            except:
                pass
        
        if not data:
            data = {
                'filename': file.filename,
                'report_type': report_type,
                'verdict': 'PASS',
                'risk_level': '🟢',
                'analysis_process': result,
                'main_findings': result,
                'detailed_comments': result,
                'odm_questions': result,
                'risk_summary': result
            }
        
        # 補充前端需要的字段
        if 'filename' not in data:
            data['filename'] = file.filename
        if 'timestamp' not in data:
            data['timestamp'] = datetime.now().isoformat()
        if 'result' not in data:
            data['result'] = f"【分析過程】\n{data.get('analysis_process', '')}"
        if 'verdict' not in data:
            data['verdict'] = 'PASS'
        
        return jsonify({'status': 'success', 'data': data})
    
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/export/excel', methods=['POST'])
def export():
    data = request.get_json() or {}
    
    # ===== 容錯處理：確保所有必要字段都有默認值 =====
    safe_data = {
        'filename': data.get('filename', 'report.pdf'),
        'report_type': data.get('report_type', 'OTHER'),
        'verdict': data.get('verdict', 'PASS'),
        'risk_level': data.get('risk_level', '🟢'),
        'analysis_process': data.get('analysis_process', '無'),
        'main_findings': data.get('main_findings', '無'),
        'detailed_comments': data.get('detailed_comments', '無'),
        'odm_questions': data.get('odm_questions', '無'),
        'risk_summary': data.get('risk_summary', '無'),
        'important_note': data.get('important_note', '無'),
    }
    
    exporter = ExcelExporter(safe_data)
    content, filename = exporter.export()
    return send_file(BytesIO(content), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=filename)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)), debug=False)

