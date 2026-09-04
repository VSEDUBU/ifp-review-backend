IFP 測試報告自動審核工具 v5.0
> 一個 RD 級通用工程審核框架，不依賴預設規則，適用於任何測試標準和供應商
![License](https://img.shields.io/badge/License-MIT-blue.svg)
![Python](https://img.shields.io/badge/Python-3.8%2B-green.svg)
![Status](https://img.shields.io/badge/Status-Production%20Ready-brightgreen.svg)
🎯 概述
這是一個為 IFP（互動式電子白板）開發團隊設計的測試報告自動審核工具。不同於傳統的"規則檢查機"，本工具採用RD 級工程審核框架：
✅ 不預設任何標準數據 - 所有限值、規範都從報告本身讀取
✅ 通用於任何標準 - EN 55032、IEC 61000、GB/T 2423、企業標準都能處理
✅ 工程師視角 - 不只判 PASS/FAIL，還識別工程風險（樣品代表性、工況覆蓋等）
✅ 即插即用 - 無需調整規則，直接上傳新報告即可審核
🚀 快速開始
需求
Python 3.8+
現代瀏覽器
安裝與運行
```bash
# 1. 克隆倉庫
git clone https://github.com/VSEDUBU/ifp-review-backend.git
cd ifp-review-backend

# 2. 安裝依賴
pip install -r requirements.txt

# 3. 配置 API 密鑰
export OPENROUTER_API_KEY="your_api_key_here"

# 4. 啟動後端
python main.py

# 5. 在瀏覽器中打開 index.html
```
📖 核心設計
RD 級審核框架（7 步）
理解報告 - 識別測試類型和採用的標準
驗證條件 - 檢查測試環境、工況、樣品代表性
檢查表格 - 驗證規格、實測值、判定的數據自洽性
覆蓋完整性 - 確認標準要求的所有項目都測了
矛盾檢測 - 封面結論與內頁數據是否一致
風險評估 - 識別工程隱患
判定結果 - PASS/FAIL/WARN/CONTRADICTION
支持的報告類型
EMI（電磁騷擾）
EMS（電磁抗擾度）
Derating（元件降額）
Safety（安規測試）
Reliability（可靠性試驗）
以及其他類型
🔑 關鍵特性
特性	說明
通用框架	不依賴硬編碼規則
智能識別	自動識別標準和限值
數據驗證	檢查數據自洽性
風險評估	工程視角的隱患識別
報告追溯	判定依據明確標注
批量處理	支持多份報告審核
多格式輸出	HTML、CSV、Excel
🔐 安全性
API 密鑰通過環境變數配置
PDF 只在內存中處理
支持 HTTPS 部署
可配置 CORS
📚 部署
本地開發
```bash
python main.py
```
Docker
```bash
docker build -t ifp-review .
docker run -e OPENROUTER_API_KEY=your_key -p 5000:5000 ifp-review
```
Heroku
```bash
heroku create ifp-review-backend
heroku config:set OPENROUTER_API_KEY=your_key
git push heroku main
```
📝 設計原則
❌ 不做
硬編碼規則數字（如"40 dBµV/m"）
預設標準數據
機械性規則應用
✅ 改為
從報告本身讀取所有標準
驗證數據的邏輯自洽性
工程師視角的風險識別
🤝 貢獻
歡迎提交 Issue 和 Pull Request。開發規範詳見 CONTRIBUTING.md
📜 許可證
MIT License
📞 聯繫
Issues: GitHub Issues
Discussion: GitHub Discussions
---
版本：5.0 | 更新時間：2026-09-04
