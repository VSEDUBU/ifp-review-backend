#!/usr/bin/env python3
"""
IFP 測試報告自動審核工具 - 後端服務 v5.0 完整版

支持 15+ 種測試報告類型，使用 RD 級通用工程審核框架
"""

import os
import json
import logging
import tempfile
from datetime import datetime
from functools import wraps
from io import BytesIO

import requests
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from dotenv import load_dotenv
import pdfplumber
import pandas as pd

# 載入環境變數
load_dotenv()

# 配置
OPENROUTER_API_KEY = os.getenv('OPENROUTER_API_KEY')
OPENROUTER_API_URL = 'https://openrouter.ai/api/v1/chat/completions'
FLASK_PORT = int(os.getenv('FLASK_PORT', 5000))
DEBUG_MODE = os.getenv('DEBUG_MODE', 'False').lower() == 'true'

# 日誌配置
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Flask 應用
app = Flask(__name__)
CORS(app)

# ─────────────────────────────────────────────────────────────────
# 支持的報告類型和對應的審核框架
# ─────────────────────────────────────────────────────────────────

REPORT_TYPES = {
    'AUTO': '🤖 自動識別',
    'EMI': 'EMI 電磁騷擾',
    'EMS': 'EMS 電磁抗擾度',
    'EMC': 'EMC 電磁相容',
    'ELEC_PERF': '電氣性能測試',
    'ELEC_TIMING': '電氣因設波測試',
    'POWER_SEQ': '電源時序測試',
    'FREQUENCY': '晶振頻偏測試',
    'SIGNAL_INT': '訊號完整性測試',
    'USB': 'USB 信號測試',
    'HDMI': 'HDMI 信號測試',
    'LVDS': 'LVDS 信號測試',
    'RELIABILITY': '可靠性試驗',
    'TEMP_RISE': '溫升試驗',
    'THERMAL': '熱測試',
    'VIBRATION': '振動測試',
    'HUMIDITY': '濕度測試',
    'ENERGY': '能效測試',
    'PERFORMANCE': '性能測試',
    'THERMAL_DESIGN': '熱設計測試',
    'DERATING': 'Derating 降額分析',
    'SAFETY': '安規測試',
    'EMI_DESIGN': 'EMI 設計測試',
    'VPC_PLATFORM': 'VPC 產品線測試',
    'SOFTWARE': '軟體/韌體測試',
    'WIRELESS': '無線模組測試',
    'OTHER': '其他測試'
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
    """調用 OpenRouter API"""
    try:
        headers = {
            'Authorization': f'Bearer {OPENROUTER_API_KEY}',
            'Content-Type': 'application/json',
            'HTTP-Referer': 'https://ifp-review-backend.onrender.com',
            'X-Title': 'IFP Test Report Review'
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
        else:
            logger.error(f"API 響應異常: {result}")
            return ""
    except Exception as e:
        logger.error(f"API 請求失敗: {str(e)}")
        return ""

def build_audit_prompt(report_type: str, stage: str) -> tuple:
    """構建審核的系統提示"""
    
    # 根據報告類型選擇對應的審核框架
    type_frameworks = {
        'EMI': '電磁騷擾測試 - 檢查輻射和傳導限值、頻段覆蓋、工況完整性',
        'EMS': '電磁抗擾度測試 - 檢查各項抗擾度等級、測試條件、判定邏輯',
        'DERATING': '降額分析 - 檢查元件應力百分比、判定準則、裕度評估',
        'RELIABILITY': '可靠性試驗 - 檢查測試條件、失效標準、樣品數量',
        'SAFETY': '安規測試 - 檢查標準符合性、測試項目完整性、隱患識別',
        'ENERGY': '能效測試 - 檢查功耗限值、測試工況、工作模式覆蓋',
        'SIGNAL_INT': '訊號完整性 - 檢查眼圖、時序參數、信噪比等',
        'USB': 'USB 信號測試 - 檢查 USB 規範符合性、眼圖、差分阻抗',
        'HDMI': 'HDMI 信號測試 - 檢查 HDMI 規範符合性、色彩深度、幀率支持',
        'ELEC_PERF': '電氣性能測試 - 檢查電壓、電流、功耗等基本參數',
        'POWER_SEQ': '電源時序測試 - 檢查上電下電順序、時序參數、邊界條件',
        'THERMAL': '熱測試 - 檢查溫升、散熱設計、工作溫度範圍',
        'FREQUENCY': '晶振頻偏測試 - 檢查頻率偏差、溫漂、工作條件',
        'PERFORMANCE': '性能測試 - 檢查處理能力、吞吐量、響應時間',
        'THERMAL_DESIGN': '熱設計測試 - 檢查熱設計功率、散熱方案、溫度均勻性',
        'VPC_PLATFORM': 'VPC 產品線測試 - 檢查平台特定的測試項目和要求',
        'SOFTWARE': '軟體/韌體測試 - 檢查版本、功能測試、穩定性測試',
        'WIRELESS': '無線模組測試 - 檢查傳輸功率、接收靈敏度、連接穩定性',
        'EMI_DESIGN': 'EMI 設計測試 - 檢查 PCB 設計、屏蔽、濾波等設計合規性',
        'EMC': '電磁相容測試 - 綜合檢查輻射、傳導、抗擾度等多個方面',
    }
    
    framework = type_frameworks.get(report_type, '通用測試 - 按標準要求進行綜合檢查')

    system_prompt = f"""你是一位資深硬體研發工程師，負責從工程視角審核 {REPORT_TYPES.get(report_type, report_type)} 報告。

【核心原則 - 重要】
⚠️  這個審核工具不使用任何硬編碼的標準限值
- 所有判定依據都必須來自報告本身的數據
- 不預設任何標準數字（如 40 dBµV/m、80% 等）
- 從報告讀取採用的標準、限值、測試條件

【你的角色】
- 工程顧問，而不是規則檢查機
- 識別工程風險和數據缺口
- 用 RD 級思維進行評估

【RD 級審核框架（7 步）】

1️⃣ 理解報告
   - 採用什麼標準？版本/年份？
   - 有無企業補充標準？

2️⃣ 驗證測試條件
   - 環境條件記錄清楚嗎？
   - 是最惡劣工況還是常規工況？
   - 樣品數量和代表性？

3️⃣ 檢查數據自洽性（關鍵）
   - 表頭清晰嗎？
   - 規格、實測、判定的邏輯是否相符？
   - 計算是否正確？

4️⃣ 檢查覆蓋完整性
   - 應測項目都測了嗎？
   - 遺漏項有合理說明嗎？

5️⃣ 矛盾檢測
   - 封面結論 vs 內頁數據是否一致？
   - 數據間是否矛盾？

6️⃣ 工程風險評估
   - 樣品代表性、邊界值、工況覆蓋、數據紧凑度

7️⃣ 最終判定
   PASS、FAIL、WARN、CONTRADICTION

【審核重點】
{framework}

【輸出格式 - 必須完全遵守】

## 報告基本資訊
- 採用標準：
- 測試工況：
- 樣品數量：

## 逐項審核
（詳細逐項檢查表）

## 風險評估
（工程視角的隱患識別）

## 矛盾偵測
（結論與數據一致性檢查）

## 整體判定
**最終判定：PASS / FAIL / WARN / CONTRADICTION**
"""

    user_prompt = f"""【待審核報告】
測試階段：{stage}
報告類型：{REPORT_TYPES.get(report_type, report_type)}

以下是報告的完整文本內容，請依照上述框架逐項審核：

---

[報告文本將在此處插入]

---

請按照 RD 級框架完成審核並輸出結果。
"""

    return system_prompt, user_prompt

# ─────────────────────────────────────────────────────────────────
# API 端點
# ─────────────────────────────────────────────────────────────────

@app.route('/health', methods=['GET'])
def health_check():
    """健康檢查"""
    return jsonify({
        'status': 'healthy',
        'version': '5.0',
        'timestamp': datetime.now().isoformat(),
        'supported_types': len(REPORT_TYPES)
    })

@app.route('/api/review', methods=['POST'])
def review_report():
    """上傳 PDF 並進行審核"""
    try:
        if not OPENROUTER_API_KEY:
            return jsonify({'status': 'error', 'message': '未設置 API 密鑰'}), 500

        if 'file' not in request.files:
            return jsonify({'status': 'error', 'message': '缺少文件'}), 400

        file = request.files['file']
        if file.filename == '' or not file.filename.endswith('.pdf'):
            return jsonify({'status': 'error', 'message': '只支持 PDF 文件'}), 400

        # 獲取參數
        report_type = request.form.get('type', 'AUTO')
        stage = request.form.get('stage', 'DVT')

        # 驗證報告類型
        if report_type not in REPORT_TYPES and report_type != 'AUTO':
            return jsonify({'status': 'error', 'message': '不支持的報告類型'}), 400

        logger.info(f"開始審核: {file.filename} | 類型: {report_type} | 階段: {stage}")

        # 提取 PDF 文本
        pdf_text = extract_pdf_text(file.stream)
        if not pdf_text:
            return jsonify({'status': 'error', 'message': 'PDF 文本提取失敗'}), 400

        # 構建提示
        system_prompt, user_prompt_template = build_audit_prompt(report_type, stage)
        user_prompt = user_prompt_template.replace('[報告文本將在此處插入]', pdf_text[:15000])

        # 調用 API
        logger.info("調用 OpenRouter API...")
        result = call_openrouter_api(system_prompt, user_prompt)

        if not result:
            return jsonify({'status': 'error', 'message': 'API 調用失敗'}), 500

        # 解析判定結果
        verdict = 'PASS'
        if 'FAIL' in result.upper() and 'CONTRADICTION' not in result.upper():
            verdict = 'FAIL'
        elif 'WARN' in result.upper():
            verdict = 'WARN'
        elif 'CONTRADICTION' in result.upper():
            verdict = 'CONTRADICTION'

        logger.info(f"審核完成: 判定={verdict}")

        return jsonify({
            'status': 'success',
            'report_id': datetime.now().strftime('%Y%m%d%H%M%S'),
            'filename': file.filename,
            'report_type': REPORT_TYPES.get(report_type, report_type),
            'stage': stage,
            'verdict': verdict,
            'result': result,
            'timestamp': datetime.now().isoformat()
        })

    except Exception as e:
        logger.error(f"審核失敗: {str(e)}")
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/review/html', methods=['POST'])
def export_html_report():
    """導出 HTML 報告"""
    try:
        data = request.get_json()
        result_text = data.get('result', '')
        filename = data.get('filename', 'report')

        html_content = f"""
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <title>{filename}</title>
    <style>
        body {{ font-family: 'Segoe UI', Arial; margin: 2cm; }}
        h1 {{ color: #0066cc; border-bottom: 2px solid #0066cc; }}
        pre {{ background: #f5f5f5; padding: 15px; border-radius: 5px; }}
        .meta {{ background: #e3f2fd; padding: 15px; border-radius: 4px; }}
    </style>
</head>
<body>
    <h1>IFP 測試報告審核結果</h1>
    <div class="meta">
        <p>生成時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        <p>報告檔名: {filename}</p>
    </div>
    <pre>{result_text}</pre>
    <p style="color: #999; font-size: 12px;">IFP 測試報告自動審核工具 v5.0</p>
</body>
</html>
        """

        buffer = BytesIO(html_content.encode('utf-8'))
        return send_file(
            buffer,
            mimetype='text/html',
            as_attachment=True,
            download_name=f'{filename}_審核報告.html'
        )
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/standards', methods=['GET'])
def get_standards():
    """獲取支持的報告類型列表"""
    return jsonify({
        'status': 'success',
        'report_types': REPORT_TYPES,
        'total': len(REPORT_TYPES)
    })

@app.route('/api/version', methods=['GET'])
def get_version():
    """獲取版本信息"""
    return jsonify({
        'version': '5.0',
        'release_date': '2026-09-04',
        'framework': 'RD-Level Engineering Audit',
        'supported_report_types': len(REPORT_TYPES),
        'features': [
            'No hardcoded rules',
            'Dynamic standard detection',
            'Engineering risk assessment',
            '15+ report types supported'
        ]
    })

if __name__ == '__main__':
    if not OPENROUTER_API_KEY:
        logger.error("❌ 未設置 OPENROUTER_API_KEY 環境變數")
        exit(1)

    logger.info(f"🚀 IFP 測試報告審核工具 v5.0 - 完整版")
    logger.info(f"📡 監聽地址: http://0.0.0.0:{FLASK_PORT}")
    logger.info(f"📊 支持 {len(REPORT_TYPES)} 種報告類型")
    logger.info("")

    app.run(host='0.0.0.0', port=FLASK_PORT, debug=DEBUG_MODE)
