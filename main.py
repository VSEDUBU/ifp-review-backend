#!/usr/bin/env python3
"""
IFP 測試報告自動審核工具 v9.1 Final
後端服務 - 手動模式 + 自動模式 + JSON 轉 Excel
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
from openpyxl.styles import Font, PatternFill, Alignment

load_dotenv()

GROQ_API_KEY = os.getenv('GROQ_API_KEY')
OPENROUTER_API_KEY = os.getenv('OPENROUTER_API_KEY')

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

UNIVERSAL_FRAMEWORK = '''【IFP 測試報告審核框架】

你是資深硬體工程師。

【RD 7 步】
1️⃣ 理解報告
2️⃣ Metadata 驗證
3️⃣ 檢查數據
4️⃣ 覆蓋完整性
5️⃣ 矛盾檢測
6️⃣ 風險評估 + 風險燈號
7️⃣ 最終判定

【輸出為 JSON】
```json
{
  "filename": "[報告名稱]",
  "report_type": "[類型]",
  "verdict": "PASS|FAIL|WARN|CONTRADICTION",
  "risk_level": "🔴|🟡|🟢",
  "analysis_process": "【分析過程】...",
  "main_findings": "【主要發現】...",
  "detailed_comments": "【詳細評論】...",
  "odm_questions": "【ODM 追問清單】...",
  "risk_summary": "【風險等級判定】...",
  "important_note": "【重要提醒】..."
}
```'''

TYPE_RULES = {
    'DERATING': '【DERATING】降額分析規則...',
    'SAFETY_IEC62368': '【SAFETY】安規測試規則...',
    'EMI': '【EMI】電磁騷擾規則...',
    'OTHER': '【OTHER】自動識別...'
}

def extract_pdf_text(pdf_file):
    try:
        text = ""
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text += page.extract_text() + "\n"
        return text
    except:
        return ""

def call_ai_api(system_prompt, user_prompt):
    if GROQ_API_KEY:
        try:
            headers = {'Authorization': f'Bearer {GROQ_API_KEY}', 'Content-Type': 'application/json'}
            data = {
                'model': 'mixtral-8x7b-32768',
                'messages': [{'role': 'system', 'content': system_prompt}, {'role': 'user', 'content': user_prompt}],
                'max_tokens': 2000
            }
            response = requests.post('https://api.groq.com/openai/v1/chat/completions', json=data, headers=headers, timeout=60)
            if response.status_code == 200:
                return response.json()['choices'][0]['message']['content']
        except:
            pass
    
    if OPENROUTER_API_KEY:
        try:
            headers = {'Authorization': f'Bearer {OPENROUTER_API_KEY}', 'Content-Type': 'application/json'}
            data = {
                'model': 'openai/gpt-3.5-turbo',
                'messages': [{'role': 'system', 'content': system_prompt}, {'role': 'user', 'content': user_prompt}],
                'max_tokens': 2000
            }
            response = requests.post('https://openrouter.ai/api/v1/chat/completions', json=data, headers=headers, timeout=60)
            if response.status_code == 200:
                return response.json()['choices'][0]['message']['content']
        except:
            pass
    
    return "❌ 所有 API 都不可用"

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
        
        sections = [
            ('報告檔名', self.data.get('filename', 'N/A')),
            ('報告類型', self.data.get('report_type', 'N/A')),
            ('最終判定', self.data.get('verdict', 'N/A')),
            ('風險燈號', self.data.get('risk_level', 'N/A')),
        ]
        
        for label, value in sections:
            ws[f'A{row}'] = label
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'B{row}'] = value
            row += 1
        
        row += 1
        
        content_sections = [
            ('分析過程', 'analysis_process'),
            ('主要發現', 'main_findings'),
            ('詳細評論', 'detailed_comments'),
            ('ODM 追問清單', 'odm_questions'),
            ('風險等級判定', 'risk_summary'),
        ]
        
        for title, key in content_sections:
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
        
        filename = f"{self.data.get('filename', 'report').replace('.pdf', '')}_審核結果.xlsx"
        return output.getvalue(), filename

@app.route('/', methods=['GET'])
def index():
    try:
        with open('index.html', 'r', encoding='utf-8') as f:
            return f.read()
    except:
        return "前端文件未找到", 500

@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({'status': 'healthy', 'version': '9.1-Final'})

@app.route('/api/generate-prompt', methods=['POST'])
def generate_prompt():
    try:
        data = request.get_json()
        report_type = data.get('type', 'OTHER')
        type_rules = TYPE_RULES.get(report_type, TYPE_RULES['OTHER'])
        prompt = f'{UNIVERSAL_FRAMEWORK}\n\n【報告類型】{report_type}\n\n{type_rules}\n\n【請粘貼報告文本】'
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
        return send_file(BytesIO(content), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=filename)
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)), debug=False)

