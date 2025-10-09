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