#!/usr/bin/env python3
"""
IFP 測試報告自動審核工具 - 後端服務 v5.0 混合模式

支持自動模式（API）和手動模式（Prompt 生成）
"""

import os
import json
import logging
from datetime import datetime

import requests
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from dotenv import load_dotenv
import pdfplumber

load_dotenv()

OPENROUTER_API_KEY = os.getenv('OPENROUTER_API_KEY')
OPENROUTER_API_URL = 'https://openrouter.ai/api/v1/chat/completions'
FLASK_PORT = int(os.getenv('FLASK_PORT', 5000))
DEBUG_MODE = os.getenv('DEBUG_MODE', 'False').lower() == 'true'

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

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
    """調用 OpenRouter API"""
    try:
        headers = {
            'Authorization': f'Bearer {OPENROUTER_API_KEY}',
            'Content-Type': 'application/json',
        }

        data = {
            'model': 'meta-llama/llama-2-70b-chat',
            'messages': [
                {'role': 'system', 'content': system_prompt},
                {'role': 'user', 'content': user_prompt}
            ],
            'temperature': 0.7,
            'max_tokens': 4000
        }

        response = requests.post(OPENROUTER_API_URL, json=data, headers=headers, timeout=60)
        response.raise_for_status()

        result = response.json()
        if 'choices' in result and len(result['choices']) > 0:
            return result['choices'][0]['message']['content']
        return ""
    except Exception as e:
        logger.error(f"API 請求失敗: {str(e)}")
        return ""

@app.route('/health', methods=['GET'])
def health_check():
    """健康檢查"""
    return jsonify({
        'status': 'healthy',
        'version': '5.0',
        'timestamp': datetime.now().isoformat()
    })

@app.route('/api/review', methods=['POST'])
def review_report():
    """上傳 PDF 並進行審核"""
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

        # 簡單的 RD 級審核框架
        system_prompt = f"""你是資深硬體工程師，審核 {report_type} 測試報告。

【RD 級審核框架】
1. 理解報告採用的標準
2. 驗證測試條件完整性
3. 檢查數據自洽性
4. 覆蓋完整性檢查
5. 矛盾偵測
6. 工程風險評估
7. 最終判定：PASS/FAIL/WARN/CONTRADICTION

【輸出格式】
## 報告基本資訊
## 逐項審核
## 風險評估
## 整體判定"""

        user_prompt = f"""測試階段：{stage}

請分析以下報告文本：
{pdf_text[:10000]}"""

        result = call_openrouter_api(system_prompt, user_prompt)

        if not result:
            return jsonify({'status': 'error', 'message': 'API 調用失敗'}), 500

        verdict = 'PASS'
        if 'FAIL' in result.upper():
            verdict = 'FAIL'
        elif 'WARN' in result.upper():
            verdict = 'WARN'
        elif 'CONTRADICTION' in result.upper():
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

@app.route('/', methods=['GET'])
def index():
    """返回前端"""
    return """
    <!DOCTYPE html>
    <html>
    <head>
        <title>IFP 審核工具</title>
        <style>
            body { font-family: Arial; margin: 40px; text-align: center; }
            h1 { color: #0066cc; }
            p { color: #666; }
        </style>
    </head>
    <body>
        <h1>🔍 IFP 測試報告審核工具 v5.0</h1>
        <p>✅ 後端服務運行中</p>
        <p><a href="/health">檢查健康狀態</a></p>
    </body>
    </html>
    """

if __name__ == '__main__':
    # 用於本地開發
    app.run(host='0.0.0.0', port=FLASK_PORT, debug=DEBUG_MODE)

# 用於生產環境（gunicorn）
# gunicorn 會直接導入 app 對象，不會執行上面的 if __name__
