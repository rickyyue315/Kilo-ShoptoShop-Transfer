# 強制指定店鋪轉貨系統

## 系統概述
基於Streamlit的零售庫存調貨建議生成系統，專為優化店鋪間庫存分配而設計。

## 主要功能
- ✅ ND/RF類型智慧識別
- ✅ 優先順序轉貨（保守/加強模式）
- ✅ 統計分析和圖表
- ✅ Excel格式匯出

## 技術棧
- 前端：Streamlit
- 資料處理：pandas, numpy
- Excel處理：openpyxl
- 視覺化：matplotlib, seaborn

## 安裝說明
```bash
pip install -r requirements.txt
```

## 運行方式
```bash
streamlit run app.py
```

## 版本資訊
- 版本：v1.0
- 開發者：Ricky

## 使用說明
1. 上傳包含庫存資料的Excel檔案
2. 選擇轉貨模式（保守轉貨或加強轉貨）
3. 點擊分析按鈕生成轉貨建議
4. 查看統計分析和圖表
5. 匯出Excel格式的轉貨建議報告

## 資料格式要求
Excel檔案必須包含以下欄位：
- Article (str) - 產品編號
- Article Description (str) - 產品描述
- RP Type (str) - 補貨類型：ND（不補貨）或 RF（補貨）
- Site (str) - 店鋪編號
- OM (str) - 營運管理單位
- MOQ (int) – 最低派貨數量
- SaSa Net Stock (int) - 現有庫存數量
- Target (int) – 目標要求數量
- Pending Received (int) - 在途訂單數量
- Safety Stock (int) - 安全庫存數量
- Last Month Sold Qty (int) - 上月銷量
- MTD Sold Qty (int) - 本月至今銷量