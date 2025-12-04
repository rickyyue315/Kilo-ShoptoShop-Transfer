# Mandatory Shop-to-Shop Transfer System

A Streamlit-based application for generating transfer recommendations between shops based on inventory data and business rules.

## Features

- **Data Upload**: Upload Excel files containing inventory and sales data
- **Data Validation**: Automatic validation of required columns and data types
- **Transfer Modes**:
  - Mode A: Conservative Transfer
  - Mode B: Enhanced Transfer
  - Mode C: Super Enhanced Transfer
- **Analytics**: Comprehensive statistics and visualizations
- **Export**: Excel export of transfer recommendations

## Transfer Mode Logic / 轉移模式邏輯

### Mode A: Conservative Transfer / 保守轉移模式

**ND 類型 (Priority 1):**
- 完整轉移所有庫存
- 條件：`SaSa Net Stock > 0`

**RF 類型 (Priority 2):**
- 條件：`(可用庫存) > Safety Stock` 且 `銷售量 < 該商品最高銷售量`
- 轉移量計算：`min(可用庫存 - Safety Stock, 可用庫存 * 0.5)`
- 最大轉移量不超過可用庫存的 **50%**
- 按銷售量升序排列（銷售較少的店舖優先轉出）

### Mode B: Enhanced Transfer / 增強轉移模式

**ND 類型 (Priority 1):**
- 完整轉移所有庫存
- 條件：`SaSa Net Stock > 0`

**RF 類型 (Priority 2):**
- 條件：`(可用庫存) > MOQ` 且 `銷售量 < 該商品最高銷售量`
- 轉移量計算：`min(可用庫存 - MOQ, 可用庫存 * 0.9)`
- 最大轉移量不超過可用庫存的 **90%**
- 按銷售量升序排列（銷售較少的店舖優先轉出）

### Mode C: Super Enhanced Transfer / 超級增強轉移模式

**ND 類型 (Priority 1):**
- 完整轉移所有庫存
- 條件：`SaSa Net Stock > 0`

**RF 類型 (Priority 2):**
- **可忽視最小庫存要求**
- 條件：`SaSa Net Stock > 0`
- 轉移量計算：`max(0, SaSa Net Stock)`
- 最大轉移量可達 **100%**，以滿足目標需求
- 按銷售量升序排列（過去銷售最多的店舖排最後出貨）

### 通用規則

- **可用庫存** = `SaSa Net Stock + Pending Received`
- **有效銷售量** = `Last Month Sold Qty` (若 > 0)，否則使用 `MTD Sold Qty`
- 轉移量不能超過實際庫存 (`SaSa Net Stock`)
- 轉移匹配時，同一 Article 和 OM 的轉出店舖與接收店舖必須不同

## Installation

1. Clone the repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the application:
   ```bash
   streamlit run app.py
   ```

## Usage

1. Upload your Excel file containing the required columns
2. Select the transfer mode (A, B, or C)
3. Click "Generate Recommendations" to process the data
4. View statistics, charts, and transfer recommendations
5. Export results to Excel

## Required Excel Columns

- Article (str) - Product code
- Article Description (str) - Product description
- RP Type (str) - Replenishment type (ND or RF)
- Site (str) - Shop code
- OM (str) - Operations manager
- MOQ (int) - Minimum order quantity
- SaSa Net Stock (int) - Current stock
- Target (int) - Target quantity
- Pending Received (int) - Pending incoming stock
- Safety Stock (int) - Safety stock level
- Last Month Sold Qty (int) - Last month's sales
- MTD Sold Qty (int) - Month-to-date sales

## Developer

Ricky

## Version

v1.0