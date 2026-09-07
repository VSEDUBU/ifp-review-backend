#!/usr/bin/env python3
"""
IFP 測試報告自動審核工具 - 企業級後端服務 v6.0 (自動模式修復版)
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

OPENROUTER_API_KEY = os.getenv('OPENROUTER_API_KEY')
OPENROUTER_API_URL = 'https://openrouter.ai/api/v1/chat/completions'
FLASK_PORT = int(os.getenv('FLASK_PORT', 5000))
DEBUG_MODE = os.getenv('DEBUG_MODE', 'False').lower() == 'true'

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

# ===== 多語言配置 =====
LANGUAGES = {
    'zh_TW': {
        'title': 'IFP 測試報告審核結果',
        'pass': '✅ 通過',
        'fail': '❌ 失敗',
        'warn': '⚠️ 警告',
        'contradiction': '🔴 矛盾',
    },
    'en_US': {
        'title': 'IFP Test Report Review Result',
        'pass': '✅ PASS',
        'fail': '❌ FAIL',
        'warn': '⚠️ WARN',
        'contradiction': '🔴 CONTRADICTION',
    }
}

# ===== 數字簽章配置 =====
SIGNATURE_CONFIG = {
    'company': 'CVTE Electronics',
    'department': 'R&D - Hardware Team',
    'certification': 'ISO 9001:2015 Certified',
    'contact': 'rd@example.com',
}

def extract_pdf_text(pdf_file) -> str:
    """從 PDF 提取文本"""
    try:
        text = ""
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text += page.extract_text() + "\n"
        return text
    except Exception as e:
        logger.error(f"PDF 提取失敗: {str(e)}")
        return ""

def call_openrouter_api(system_prompt: str, user_prompt: str) -> str:
    """調用 OpenRouter API - 修復版"""
    try:
        if not OPENROUTER_API_KEY:
            return "❌ 錯誤：OPENROUTER_API_KEY 未設置。請在 Render 環境變數中設置。"
        
        # 使用通用的 gpt-3.5-turbo 模型（更穩定）
        headers = {
            'Authorization': f'Bearer {OPENROUTER_API_KEY}',
            'Content-Type': 'application/json',
            'HTTP-Referer': 'https://ifp-review-backend.onrender.com',
            'X-Title': 'IFP Test Report Review Tool'
        }

        data = {
            'model': 'openai/gpt-3.5-turbo',  # 改用通用模型
            'messages': [
                {'role': 'system', 'content': system_prompt},
                {'role': 'user', 'content': user_prompt}
            ],
            'temperature': 0.7,
            'max_tokens': 2000
        }

        logger.info(f"調用 API: {data['model']}")
        
        response = requests.post(
            OPENROUTER_API_URL, 
            json=data, 
            headers=headers, 
            timeout=60
        )
        
        logger.info(f"API 響應狀態碼: {response.status_code}")
        
        if response.status_code == 404:
            # 如果 gpt-3.5-turbo 不可用，嘗試 claude-3-haiku
            logger.warning("gpt-3.5-turbo 不可用，嘗試 claude-3-haiku")
            data['model'] = 'anthropic/claude-3-haiku'
            response = requests.post(
                OPENROUTER_API_URL, 
                json=data, 
                headers=headers, 
                timeout=60
            )
        
        response.raise_for_status()

        result = response.json()
        if 'choices' in result and len(result['choices']) > 0:
            return result['choices'][0]['message']['content']
        
        logger.error(f"API 返回空結果: {result}")
        return "❌ API 返回空結果"
        
    except requests.exceptions.HTTPError as e:
        logger.error(f"HTTP 錯誤: {e.response.status_code} - {e.response.text}")
        return f"❌ HTTP 錯誤: {e.response.status_code}"
    except Exception as e:
        logger.error(f"API 調用失敗: {str(e)}")
        return f"❌ API 調用失敗: {str(e)}"

# ===== Excel 導出 =====
class ExcelExporter:
    def __init__(self, data: dict):
        self.data = data

    def export(self) -> tuple:
        """導出為 Excel 格式"""
        wb = Workbook()
        ws = wb.active
        ws.title = "審核結果"
        
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 70
        
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        title_font = Font(bold=True, size=16, color="1F4E78")
        section_font = Font(bold=True, size=12, color="FFFFFF")
        section_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        
        border = Border(
            left=Side(style='thin', color="000000"),
            right=Side(style='thin', color="000000"),
            top=Side(style='thin', color="000000"),
            bottom=Side(style='thin', color="000000")
        )
        
        row = 1
        
        # 標題
        title_cell = ws[f'A{row}']
        title_cell.value = LANGUAGES['zh_TW']['title']
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
            ('項目階段', self.data.get('stage', 'N/A')),
            ('審核時間', self.data.get('timestamp', 'N/A')),
        ]
        
        for label, value in info_data:
            ws[f'A{row}'] = label
            ws[f'A{row}'].font = Font(bold=True, size=11)
            ws[f'A{row}'].fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
            ws[f'A{row}'].border = border
            
            ws[f'B{row}'] = value
            ws[f'B{row}'].border = border
            ws[f'B{row}'].alignment = Alignment(wrap_text=True, vertical='top')
            row += 1
        
        row += 1
        
        # 審核結果
        verdict = self.data.get('verdict', 'UNKNOWN')
        if verdict == 'PASS':
            verdict_color = "70AD47"
        elif verdict == 'FAIL':
            verdict_color = "FF0000"
        elif verdict == 'WARN':
            verdict_color = "FFC000"
        else:
            verdict_color = "7F7F7F"
        
        result_section = ws[f'A{row}']
        result_section.value = "📊 審核結果"
        result_section.font = section_font
        result_section.fill = section_fill
        ws.merge_cells(f'A{row}:B{row}')
        row += 1
        
        ws[f'A{row}'] = "最終判定"
        ws[f'A{row}'].font = Font(bold=True, size=11)
        ws[f'A{row}'].fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
        ws[f'A{row}'].border = border
        
        verdict_map = {
            'PASS': LANGUAGES['zh_TW']['pass'],
            'FAIL': LANGUAGES['zh_TW']['fail'],
            'WARN': LANGUAGES['zh_TW']['warn'],
            'CONTRADICTION': LANGUAGES['zh_TW']['contradiction'],
        }
        
        ws[f'B{row}'] = verdict_map.get(verdict, verdict)
        ws[f'B{row}'].font = Font(bold=True, size=12, color="FFFFFF")
        ws[f'B{row}'].fill = PatternFill(start_color=verdict_color, end_color=verdict_color, fill_type="solid")
        ws[f'B{row}'].border = border
        row += 1
        
        row += 1
        
        # 詳細分析
        analysis_section = ws[f'A{row}']
        analysis_section.value = "📝 詳細分析"
        analysis_section.font = section_font
        analysis_section.fill = section_fill
        ws.merge_cells(f'A{row}:B{row}')
        row += 1
        
        result_text = self.data.get('result', '無')
        ws[f'A{row}'] = result_text
        ws[f'A{row}'].alignment = Alignment(wrap_text=True, vertical='top')
        ws[f'A{row}'].border = border
        ws.merge_cells(f'A{row}:B{row}')
        ws.row_dimensions[row].height = 250
        
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        filename = f"{self.data.get('filename', 'report').replace('.pdf', '')}_審核結果.xlsx"
        return output.getvalue(), filename

# ===== API 端點 =====

@app.route('/', methods=['GET'])
def index():
    """返回前端工具"""
    try:
        with open('index.html', 'r', encoding='utf-8') as f:
            return f.read()
    except FileNotFoundError:
        return "前端文件未找到", 500

@app.route('/health', methods=['GET'])
def health_check():
    """健康檢查"""
    has_api_key = bool(OPENROUTER_API_KEY)
    return jsonify({
        'status': 'healthy' if has_api_key else 'warning',
        'version': '6.0',
        'api_key_configured': has_api_key,
        'timestamp': datetime.now().isoformat()
    })

@app.route('/api/review', methods=['POST'])
def review_report():
    """上傳 PDF 並進行自動審核"""
    try:
        if 'file' not in request.files:
            return jsonify({'status': 'error', 'message': '缺少文件'}), 400

        file = request.files['file']
        if file.filename == '' or not file.filename.endswith('.pdf'):
            return jsonify({'status': 'error', 'message': '只支持 PDF 文件'}), 400

        report_type = request.form.get('type', 'AUTO')
        stage = request.form.get('stage', 'DVT')

        logger.info(f"開始審核: {file.filename}")

        pdf_text = extract_pdf_text(file.stream)
        if not pdf_text:
            return jsonify({'status': 'error', 'message': 'PDF 文本提取失敗'}), 400

        # RD 級審核框架
        system_prompt = f"""你是資深硬體工程師，負責審核 {report_type} 測試報告。

【RD 級審核框架 - 7 步】
1️⃣ 理解報告 - 標準、版本、補充規範
2️⃣ 驗證條件 - 環境、工況、樣品代表性
3️⃣ 檢查數據 - 規格、實測、判定邏輯
4️⃣ 覆蓋完整性 - 應測項目是否都測了
5️⃣ 矛盾檢測 - 結論與數據一致性
6️⃣ 風險評估 - 工程隱患識別
7️⃣ 最終判定 - PASS / FAIL / WARN / CONTRADICTION

請簡潔地分析報告。"""

        user_prompt = f"""測試階段：{stage}
報告類型：{report_type}

請分析以下報告文本（前 8000 字符）：

{pdf_text[:8000]}

在最後清楚地說明：最終判定：PASS 或 FAIL 或 WARN 或 CONTRADICTION"""

        result = call_openrouter_api(system_prompt, user_prompt)

        if result.startswith("❌"):
            return jsonify({'status': 'error', 'message': result}), 500

        # 自動判定
        verdict = 'PASS'
        upper_result = result.upper()
        if 'FAIL' in upper_result and 'CONTRADICTION' not in upper_result:
            verdict = 'FAIL'
        elif 'WARN' in upper_result:
            verdict = 'WARN'
        elif 'CONTRADICTION' in upper_result:
            verdict = 'CONTRADICTION'

        return jsonify({
            'status': 'success',
            'filename': file.filename,
            'report_type': report_type,
            'stage': stage,
            'verdict': verdict,
            'result': result,
            'timestamp': datetime.now().isoformat()
        })

    except Exception as e:
        logger.error(f"審核失敗: {str(e)}")
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/generate-prompt', methods=['POST'])
def generate_prompt():
    """生成手動模式的 Prompt"""
    try:
        data = request.get_json()
        report_type = data.get('type', 'AUTO')
        stage = data.get('stage', 'DVT')
        
        prompt = f"""【IFP 測試報告 RD 級審核】

你是資深硬體工程師，負責審核 {report_type} 測試報告。

【RD 級審核框架 - 7 步】
1️⃣ 理解報告 - 標準、版本、補充規範
2️⃣ 驗證條件 - 環境、工況、樣品代表性
3️⃣ 檢查數據 - 規格、實測、判定邏輯
4️⃣ 覆蓋完整性 - 應測項目是否都測了
5️⃣ 矛盾檢測 - 結論與數據一致性
6️⃣ 風險評估 - 工程隱患識別
7️⃣ 最終判定 - PASS / FAIL / WARN / CONTRADICTION

【報告信息】
- 測試階段：{stage}
- 報告類型：{report_type}
- 審核時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

【你的任務】
請詳細分析下面的測試報告，按照 RD 級框架進行審核。

在分析完後，在最後清楚地寫出：
**最終判定：PASS / FAIL / WARN / CONTRADICTION**

【請在下方粘貼報告文本】
"""
        
        return jsonify({
            'status': 'success',
            'prompt': prompt
        })
    except Exception as e:
        logger.error(f"生成 Prompt 失敗: {str(e)}")
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/export/excel', methods=['POST'])
def export_excel():
    """導出為 Excel 格式"""
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
        logger.error(f"Excel 導出失敗: {str(e)}")
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/export/html', methods=['POST'])
def export_html():
    """導出為 HTML 格式"""
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({'status': 'error', 'message': '缺少數據'}), 400
        
        verdict = data.get('verdict', 'UNKNOWN')
        verdict_text = {
            'PASS': '✅ 通過',
            'FAIL': '❌ 失敗',
            'WARN': '⚠️ 警告',
            'CONTRADICTION': '🔴 矛盾',
        }.get(verdict, verdict)
        
        html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>IFP 測試報告審核結果</title>
            <style>
                body {{ font-family: Arial; padding: 20px; }}
                .header {{ text-align: center; border-bottom: 3px solid #1F4E78; padding-bottom: 20px; margin-bottom: 30px; }}
                .verdict {{ font-size: 24px; font-weight: bold; margin: 20px 0; }}
                .info {{ margin: 10px 0; }}
                .label {{ font-weight: bold; width: 150px; display: inline-block; }}
                .analysis {{ background: #f9f9f9; padding: 20px; border-radius: 8px; white-space: pre-wrap; margin-top: 20px; }}
            </style>
        </head>
        <body>
            <div class="header">
                <h1>IFP 測試報告審核結果</h1>
                <p>Professional Test Report Review - {datetime.now().strftime('%Y-%m-%d')}</p>
            </div>
            <div>
                <div class="info"><span class="label">報告檔名：</span>{data.get('filename', 'N/A')}</div>
                <div class="info"><span class="label">報告類型：</span>{data.get('report_type', 'N/A')}</div>
                <div class="info"><span class="label">項目階段：</span>{data.get('stage', 'N/A')}</div>
                <div class="info"><span class="label">審核時間：</span>{data.get('timestamp', 'N/A')}</div>
            </div>
            <div class="verdict">{verdict_text}</div>
            <div class="analysis">{data.get('result', '無')}</div>
        </body>
        </html>
        """
        
        filename = f"{data.get('filename', 'report').replace('.pdf', '')}_審核結果.html"
        return send_file(
            BytesIO(html.encode('utf-8')),
            mimetype='text/html; charset=utf-8',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        logger.error(f"HTML 導出失敗: {str(e)}")
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/version', methods=['GET'])
def get_version():
    """獲取版本信息"""
    return jsonify({
        'version': '6.0',
        'tier': 'Enterprise',
    })

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=FLASK_PORT, debug=DEBUG_MODE)

