# 店舖間強制轉移系統

基於 Streamlit 開發的店舖間庫存轉移建議系統，可根據庫存數據和業務規則生成智能轉移建議。

## 系統特色

- **資料上傳**：支援 Excel 檔案上傳，包含庫存和銷售數據
- **資料驗證**：自動驗證必要欄位和資料類型
- **三種轉移模式**：
  - A 模式：保守轉移
  - B 模式：增強轉移
  - C 模式：超級增強轉移
- **統計分析**：提供全面的統計數據和視覺化圖表
- **Excel 匯出**：一鍵匯出轉移建議結果

## 轉移模式邏輯

### A 模式：保守轉移

**ND 類型 (優先級 1):**
- 完整轉移所有庫存
- 條件：`SaSa Net Stock > 0`

**RF 類型 (優先級 2):**
- 條件：`(可用庫存) > Safety Stock` 且 `銷售量 < 該商品最高銷售量`
- 轉移量計算：`min(可用庫存 - Safety Stock, 可用庫存 * 0.5)`
- 最大轉移量不超過可用庫存的 **50%**
- 按銷售量升序排列（銷售較少的店舖優先轉出）
- **OM 限制**：僅限相同 OM 組別內調撥

### B 模式：增強轉移

**ND 類型 (優先級 1):**
- 完整轉移所有庫存
- 條件：`SaSa Net Stock > 0`

**RF 類型 (優先級 2):**
- 條件：`(可用庫存) > MOQ` 且 `銷售量 < 該商品最高銷售量`
- 轉移量計算：`min(可用庫存 - MOQ, 可用庫存 * 0.9)`
- 最大轉移量不超過可用庫存的 **90%**
- 按銷售量升序排列（銷售較少的店舖優先轉出）
- **OM 限制**：僅限相同 OM 組別內調撥

### C 模式：超級增強轉移

**ND 類型 (優先級 1):**
- 完整轉移所有庫存
- 條件：`SaSa Net Stock > 0`

**RF 類型 (優先級 2):**
- **可忽視最小庫存要求**
- 條件：`SaSa Net Stock > 0`
- 轉移量計算：`max(0, SaSa Net Stock)`
- 最大轉移量可達 **100%**，以滿足目標需求
- 按銷售量升序排列（過去銷售最多的店舖排最後出貨）
- **OM 限制**：**允許跨 OM 組別調撥**，但限制 **HD 組別不能調撥至 HA、HB、HC 組別**

### 通用規則

- **可用庫存** = `SaSa Net Stock + Pending Received`
- **有效銷售量** = `Last Month Sold Qty` (若 > 0)，否則使用 `MTD Sold Qty`
- 轉移量不能超過實際庫存 (`SaSa Net Stock`)
- 轉移匹配時，同一 Article 的轉出店舖與接收店舖必須不同
- **總需求限制**：總轉移量不超過該商品的總需求量

## 安裝步驟

1. 複製專案
2. 安裝依賴套件：
   ```bash
   pip install -r requirements.txt
   ```
3. 執行應用程式：
   ```bash
   streamlit run app.py
   ```

## 使用說明

1. 上傳包含必要欄位的 Excel 檔案
2. 選擇轉移模式（A、B 或 C）
3. 點擊「生成轉移建議」處理數據
4. 查看統計數據、圖表和轉移建議明細
5. 匯出結果至 Excel

## Excel 必要欄位

| 欄位名稱 | 類型 | 說明 |
|---------|------|------|
| Article | str | 商品代碼 |
| Article Description | str | 商品描述 |
| RP Type | str | 補貨類型 (ND 或 RF) |
| Site | str | 店舖代碼 |
| OM | str | 營運經理組別 |
| MOQ | int | 最小訂購量 |
| SaSa Net Stock | int | 現有庫存 |
| Target | int | 目標數量 |
| Pending Received | int | 待入庫數量 |
| Safety Stock | int | 安全庫存水平 |
| Last Month Sold Qty | int | 上月銷量 |
| MTD Sold Qty | int | 本月累計銷量 |

## Excel 輸出格式

### 轉移建議工作表

系統會生成包含以下欄位的 Excel 檔案：

1. **基本資訊**
   - Article - 商品代碼
   - Article Description - 商品描述
   - OM - 營運經理組別

2. **調撥店舖資訊**
   - Transfer Site - 調撥店舖
   - Transfer Qty - 調撥數量
   - Transfer Site Original Stock - 調撥店原始庫存
   - Transfer Site After Transfer Stock - 調撥後庫存
   - Transfer Site Safety Stock - 調撥店安全庫存
   - Transfer Site MOQ - 調撥店最小訂購量
   - Transfer Site RP Type - 調撥店 RP 類型
   - **Transfer Site Last Month Sold Qty** - 調撥店上月銷量
   - **Transfer Site MTD Sold Qty** - 調撥店本月銷量

3. **接收店舖資訊**
   - Receive Site - 接收店舖
   - Receive Site Target Qty - 接收店目標數量
   - Receive Site RP Type - 接收店 RP 類型
   - **Receive Site Last Month Sold Qty** - 接收店上月銷量
   - **Receive Site MTD Sold Qty** - 接收店本月銷量

4. **轉移資訊**
   - Transfer Type - 轉移類型 (ND Transfer / RF Excess Transfer / RF Enhanced Transfer / RF Super Enhanced Transfer)
   - Receive Qty - 實際接收數量
   - Notes - 備註

### 統計摘要工作表

包含以下統計數據：
- **基本 KPI**：總建議數、總轉移量、唯一商品數、唯一 OM 數
- **按商品統計**：商品代碼、總需求量、總轉移量、轉移行數、滿足率 (%)
- **按 OM 統計**：OM 組別、總轉移量、總需求量、轉移行數、唯一商品數
- **轉移類型分佈**：轉移類型、總量、行數
- **接收統計**：店舖代碼、總目標量、總接收量

## 開發者

Ricky

## 版本

v2.0

### 版本 2.0 更新內容

- ✅ 新增 RP Type 欄位至 Excel 輸出
- ✅ 優化 C 模式：允許跨 OM 組別調撥（HD 除外）
- ✅ 界面全面中文化
- ✅ 新增銷售數據欄位（Last Month Sold Qty, MTD Sold Qty）
- ✅ 修正總需求計算邏輯