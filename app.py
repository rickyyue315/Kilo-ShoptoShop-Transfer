import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import io
import xlsxwriter
import warnings
warnings.filterwarnings('ignore')

# Set page configuration
st.set_page_config(
    page_title="強制指定店鋪轉貨系統",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Define required columns
REQUIRED_COLUMNS = [
    'Article', 'Article Description', 'RP Type', 'Site', 'OM', 'MOQ',
    'SaSa Net Stock', 'Target', 'Pending Received', 'Safety Stock',
    'Last Month Sold Qty', 'MTD Sold Qty'
]

def validate_file_structure(df):
    """Validate that the uploaded file has all required columns"""
    missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_columns:
        return False, f"缺少必要欄位: {', '.join(missing_columns)}"
    return True, "檔案結構驗證通過"

def preprocess_data(df):
    """Preprocess the data according to business rules"""
    # Create a copy to avoid modifying original data
    processed_df = df.copy()
    
    # Initialize Notes column for data cleaning logs
    processed_df['Notes'] = ''
    
    # 1. Convert Article to string type
    processed_df['Article'] = processed_df['Article'].astype(str)
    
    # 2. Convert numeric columns to integers, fill invalid values with 0
    numeric_columns = [
        'MOQ', 'SaSa Net Stock', 'Target', 'Pending Received', 
        'Safety Stock', 'Last Month Sold Qty', 'MTD Sold Qty'
    ]
    
    for col in numeric_columns:
        # Convert to numeric, errors to NaN, then fill NaN with 0, then convert to int
        processed_df[col] = pd.to_numeric(processed_df[col], errors='coerce').fillna(0).astype(int)
    
    # 3. Correct negative inventory and sales values to 0
    processed_df['SaSa Net Stock'] = processed_df['SaSa Net Stock'].clip(lower=0)
    processed_df['Last Month Sold Qty'] = processed_df['Last Month Sold Qty'].clip(lower=0)
    processed_df['MTD Sold Qty'] = processed_df['MTD Sold Qty'].clip(lower=0)
    
    # 4. Limit sales outliers (>100000) to 100000 and add notes
    sales_outlier_mask = (processed_df['Last Month Sold Qty'] > 100000) | (processed_df['MTD Sold Qty'] > 100000)
    processed_df.loc[sales_outlier_mask, 'Notes'] = '銷量異常值已限制為100000'
    processed_df['Last Month Sold Qty'] = processed_df['Last Month Sold Qty'].clip(upper=100000)
    processed_df['MTD Sold Qty'] = processed_df['MTD Sold Qty'].clip(upper=100000)
    
    # 5. Fill string columns with empty string
    string_columns = ['Article Description', 'RP Type', 'Site', 'OM']
    for col in string_columns:
        processed_df[col] = processed_df[col].fillna('').astype(str)
    
    # 6. Validate RP Type values
    invalid_rp_mask = ~processed_df['RP Type'].isin(['ND', 'RF'])
    processed_df.loc[invalid_rp_mask, 'Notes'] += ' RP Type無效，已設為ND'
    processed_df.loc[invalid_rp_mask, 'RP Type'] = 'ND'
    
    return processed_df

def calculate_effective_sales(row):
    """Calculate effective sales based on business rules"""
    if row['Last Month Sold Qty'] > 0:
        return row['Last Month Sold Qty']
    else:
        return row['MTD Sold Qty']

def identify_transfer_out_candidates(df, mode='conservative'):
    """Identify transfer-out candidates based on selected mode"""
    transfer_out_candidates = []

    # First, calculate total demand for each article across all OMs
    total_demand_by_article = df[df['Target'] > 0].groupby('Article')['Target'].sum()

    # Group by Article and OM for processing
    grouped = df.groupby(['Article', 'OM'])

    for (article, om), group in grouped:
        # Calculate effective sales for each store
        group['Effective Sales'] = group.apply(calculate_effective_sales, axis=1)

        # Calculate max sales for this article across all stores in this OM
        max_sales = group['Effective Sales'].max()

        # Get total demand for this article across all OMs
        article_total_demand = total_demand_by_article.get(article, 0)

        # Priority 1: ND Type Complete Transfer Out
        nd_stores = group[group['RP Type'] == 'ND']
        for _, store in nd_stores.iterrows():
            if store['SaSa Net Stock'] > 0:
                transfer_out_candidates.append({
                    'Article': article,
                    'OM': om,
                    'Site': store['Site'],
                    'Transfer Type': 'ND轉出',
                    'Available Stock': store['SaSa Net Stock'],
                    'Transfer Qty': store['SaSa Net Stock'],
                    'Effective Sales': store['Effective Sales'],
                    'Original Stock': store['SaSa Net Stock'],
                    'Safety Stock': store['Safety Stock'],
                    'MOQ': store['MOQ'],
                    'Pending Received': store['Pending Received'],
                    'Article Total Demand': article_total_demand
                })

        # Priority 2: RF Type Transfer Out (different logic for conservative vs enhanced)
        rf_stores = group[group['RP Type'] == 'RF']

        # Sort by effective sales (lowest first for transfer-out priority)
        rf_stores_sorted = rf_stores.sort_values('Effective Sales', ascending=True)

        for _, store in rf_stores_sorted.iterrows():
            total_available = store['SaSa Net Stock'] + store['Pending Received']

            if mode == 'conservative':
                # Conservative approach
                if total_available > store['Safety Stock']:
                    base_transferable = total_available - store['Safety Stock']
                    max_transferable = int(total_available * 0.5)
                    actual_transfer = min(base_transferable, max_transferable, store['SaSa Net Stock'])

                    if actual_transfer > 0:
                        transfer_out_candidates.append({
                            'Article': article,
                            'OM': om,
                            'Site': store['Site'],
                            'Transfer Type': 'RF過剩轉出',
                            'Available Stock': store['SaSa Net Stock'],
                            'Transfer Qty': actual_transfer,
                            'Effective Sales': store['Effective Sales'],
                            'Original Stock': store['SaSa Net Stock'],
                            'Safety Stock': store['Safety Stock'],
                            'MOQ': store['MOQ'],
                            'Pending Received': store['Pending Received'],
                            'Article Total Demand': article_total_demand
                        })

            elif mode == 'enhanced':
                # Enhanced approach
                if total_available > (store['MOQ'] + 1):
                    base_transferable = total_available - (store['MOQ'] + 1)
                    max_transferable = int(total_available * 0.8)
                    actual_transfer = min(base_transferable, max_transferable, store['SaSa Net Stock'])

                    if actual_transfer > 0:
                        transfer_out_candidates.append({
                            'Article': article,
                            'OM': om,
                            'Site': store['Site'],
                            'Transfer Type': 'RF加強轉出',
                            'Available Stock': store['SaSa Net Stock'],
                            'Transfer Qty': actual_transfer,
                            'Effective Sales': store['Effective Sales'],
                            'Original Stock': store['SaSa Net Stock'],
                            'Safety Stock': store['Safety Stock'],
                            'MOQ': store['MOQ'],
                            'Pending Received': store['Pending Received'],
                            'Article Total Demand': article_total_demand
                        })

            else:  # special enhanced mode
                # Special Enhanced approach (特強轉貨)
                if total_available > 0 and store['Effective Sales'] < max_sales:
                    # Base transferable = Transfer Qty - (Stock + In Transit) leaving 2 units
                    base_transferable = store['SaSa Net Stock'] - 2
                    max_transferable = int(total_available * 0.9)
                    actual_transfer = min(base_transferable, max_transferable, store['SaSa Net Stock'])

                    if actual_transfer > 0:
                        transfer_out_candidates.append({
                            'Article': article,
                            'OM': om,
                            'Site': store['Site'],
                            'Transfer Type': 'RF特強轉出',
                            'Available Stock': store['SaSa Net Stock'],
                            'Transfer Qty': actual_transfer,
                            'Effective Sales': store['Effective Sales'],
                            'Original Stock': store['SaSa Net Stock'],
                            'Safety Stock': store['Safety Stock'],
                            'MOQ': store['MOQ'],
                            'Pending Received': store['Pending Received'],
                            'Article Total Demand': article_total_demand
                        })

    return pd.DataFrame(transfer_out_candidates)

def identify_transfer_in_candidates(df):
    """Identify transfer-in candidates (stores with Target values)"""
    transfer_in_candidates = []
    
    # Filter stores with Target > 0
    target_stores = df[df['Target'] > 0]
    
    for _, store in target_stores.iterrows():
        transfer_in_candidates.append({
            'Article': store['Article'],
            'OM': store['OM'],
            'Site': store['Site'],
            'Transfer Type': '指定店鋪補貨',
            'Required Qty': store['Target'],
            'Effective Sales': calculate_effective_sales(store),
            'Current Stock': store['SaSa Net Stock'],
            'Safety Stock': store['Safety Stock'],
            'MOQ': store['MOQ']
        })
    
    return pd.DataFrame(transfer_in_candidates)

def handle_no_transfer_candidates(transfer_out_df, transfer_in_df, mode):
    """Handle scenario when no eligible transfer candidates are found"""
    import logging

    # Configure logging
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger(__name__)

    # Analyze the situation
    no_out_candidates = transfer_out_df.empty
    no_in_candidates = transfer_in_df.empty

    # Create diagnostic information
    diagnostic_info = {
        'mode': mode,
        'transfer_out_count': len(transfer_out_df),
        'transfer_in_count': len(transfer_in_df),
        'reason': 'unknown'
    }

    if no_out_candidates and no_in_candidates:
        diagnostic_info['reason'] = 'no_eligible_candidates'
        message = "沒有找到符合轉出或轉入條件的候選商店。請檢查資料是否包含：\n" \
                 "• ND類型且庫存大於0的產品\n" \
                 "• 具有目標需求量的產品"
    elif no_out_candidates:
        diagnostic_info['reason'] = 'no_transfer_out_candidates'
        message = "沒有找到符合轉出條件的候選商店。請檢查：\n" \
                 "• 是否有ND類型產品且庫存大於0\n" \
                 "• RF類型產品是否滿足轉出條件（依所選模式而定）"
    elif no_in_candidates:
        diagnostic_info['reason'] = 'no_transfer_in_candidates'
        message = "沒有找到符合轉入條件的候選商店。請檢查：\n" \
                 "• 是否有產品設置了目標需求量（Target > 0）"
    else:
        # Check for common articles between out and in candidates
        out_articles = set(transfer_out_df['Article'].unique())
        in_articles = set(transfer_in_df['Article'].unique())
        common_articles = out_articles.intersection(in_articles)

        if not common_articles:
            diagnostic_info['reason'] = 'no_common_articles'
            message = "沒有找到可以匹配的產品。轉出候選和轉入候選的產品沒有重疊。"
        else:
            diagnostic_info['reason'] = 'om_constraint_violation'
            message = "沒有找到符合OM約束的轉貨機會。系統要求轉出和轉入必須在同一OM單位內。"

    # Log diagnostic information
    logger.info(f"No transfer suggestions generated for mode: {mode}")
    logger.info(f"Diagnostic info: {diagnostic_info}")

    # Create user-friendly error response
    error_response = {
        'success': False,
        'message': message,
        'diagnostic': diagnostic_info,
        'suggestions': [
            "檢查Excel檔案是否包含所有必要欄位",
            "確認是否有ND類型產品且庫存大於0",
            "確認是否有產品設置了目標需求量",
            "檢查轉出和轉入產品是否屬於同一OM單位",
            "驗證資料格式是否正確（數值欄位應為數字）"
        ]
    }

    return error_response

def match_transfers(transfer_out_df, transfer_in_df, original_df):
    """Match transfer-out and transfer-in candidates with proper demand constraint"""
    transfer_suggestions = []

    # Check if either dataframe is empty
    if transfer_out_df.empty or transfer_in_df.empty:
        return pd.DataFrame(transfer_suggestions)

    # Make a copy of transfer_in_df to avoid modifying original
    transfer_in_df_copy = transfer_in_df.copy()

    # Group by Article to apply constraint at article level
    out_grouped = transfer_out_df.groupby(['Article'])
    in_grouped = transfer_in_df_copy.groupby(['Article'])

    for article, out_group in out_grouped:
        if article in in_grouped.groups:
            in_group = in_grouped.get_group(article)

            # Calculate total demand for this article across all OMs
            total_demand = in_group['Required Qty'].sum()

            # Get all transfer-out candidates for this article across all OMs
            out_group_sorted = out_group.sort_values(['OM', 'Transfer Type', 'Effective Sales'],
                                                   ascending=[True, True, True])

            # Sort transfer-in by OM, then by Effective Sales (highest first)
            in_group_sorted = in_group.sort_values(['OM', 'Effective Sales'], ascending=[True, False])

            # Track total transferred for this article across all OMs
            total_transferred = 0

            # Match transfers with proper constraint enforcement
            for _, out_store in out_group_sorted.iterrows():
                remaining_qty = out_store['Transfer Qty']

                for idx, in_store in in_group_sorted.iterrows():
                    if remaining_qty <= 0:
                        break

                    # Avoid self-transfer
                    if out_store['Site'] == in_store['Site']:
                        continue

                    # Calculate potential transfer quantity
                    potential_transfer_qty = min(remaining_qty, in_store['Required Qty'])

                    # Apply global demand constraint for this article
                    if total_transferred + potential_transfer_qty > total_demand:
                        potential_transfer_qty = max(0, total_demand - total_transferred)

                    if potential_transfer_qty > 0:
                        # Get product description from original data
                        product_desc = original_df[original_df['Article'] == article]['Article Description'].iloc[0]

                        transfer_suggestions.append({
                            'Article': article,
                            'Product Desc': product_desc,
                            'OM': out_store['OM'],
                            'Transfer Site': out_store['Site'],
                            'Transfer Qty': potential_transfer_qty,
                            'Transfer Site Original Stock': out_store['Original Stock'],
                            'Transfer Site After Transfer Stock': out_store['Original Stock'] - potential_transfer_qty,
                            'Transfer Site Safety Stock': out_store['Safety Stock'],
                            'Transfer Site MOQ': out_store['MOQ'],
                            'Receive Site': in_store['Site'],
                            'Receive Site Target Qty': in_store['Required Qty'],
                            'Transfer Type': out_store['Transfer Type'],
                            'Notes': f"從{out_store['Site']}轉移至{in_store['Site']}"
                        })

                        # Update tracking variables
                        remaining_qty -= potential_transfer_qty
                        total_transferred += potential_transfer_qty

                        # Update the required quantity for the receiving store (in copy)
                        transfer_in_df_copy.loc[idx, 'Required Qty'] -= potential_transfer_qty

                        # Update the sorted in_group for next iteration
                        in_group_sorted.loc[idx, 'Required Qty'] -= potential_transfer_qty

    return pd.DataFrame(transfer_suggestions)

def calculate_statistics(transfer_suggestions_df, mode):
    """Calculate comprehensive statistics with demand constraint validation"""
    if transfer_suggestions_df.empty:
        return {
            'total_transfer_qty': 0,
            'total_transfer_lines': 0,
            'unique_articles': 0,
            'unique_oms': 0,
            'article_stats': pd.DataFrame(),
            'om_stats': pd.DataFrame(),
            'transfer_type_stats': pd.DataFrame(),
            'receive_stats': pd.DataFrame(),
            'constraint_violations': 0,
            'violation_details': []
        }

    # Basic KPIs
    total_transfer_qty = transfer_suggestions_df['Transfer Qty'].sum()
    total_transfer_lines = len(transfer_suggestions_df)
    unique_articles = transfer_suggestions_df['Article'].nunique()
    unique_oms = transfer_suggestions_df['OM'].nunique()

    # Calculate total demand by Article for constraint validation
    total_demand_by_article = transfer_suggestions_df.groupby('Article')['Receive Site Target Qty'].first()
    total_transfer_by_article = transfer_suggestions_df.groupby('Article')['Transfer Qty'].sum()

    # Check for constraint violations
    constraint_violations = 0
    violation_details = []

    for article in total_demand_by_article.index:
        demand = total_demand_by_article[article]
        transfer = total_transfer_by_article.get(article, 0)
        if transfer > demand:
            constraint_violations += 1
            violation_details.append({
                'Article': article,
                'Total Demand': demand,
                'Total Transfer': transfer,
                'Violation': transfer - demand
            })

    # Statistics by Article (updated with constraint validation)
    article_stats = transfer_suggestions_df.groupby('Article').agg({
        'Receive Site Target Qty': 'first',  # Total demand
        'Transfer Qty': 'sum',  # Total transfer
        'OM': 'nunique'  # Number of OMs involved
    }).round(2)
    article_stats.columns = ['總需求件數', '總調貨件數', '涉及OM數量']
    article_stats['轉貨行數'] = transfer_suggestions_df.groupby('Article').size()
    article_stats['需求滿足率'] = (article_stats['總調貨件數'] / article_stats['總需求件數'] * 100).round(2)
    article_stats['約束違規'] = [(total_transfer_by_article.get(article, 0) > article_stats.loc[article, '總需求件數']) for article in article_stats.index]
    
    # Statistics by OM (updated with new requirements)
    om_stats = transfer_suggestions_df.groupby('OM').agg({
        'Receive Site Target Qty': 'first',  # Total demand
        'Transfer Qty': 'sum',  # Total transfer
        'Article': 'nunique'  # Number of products involved
    }).round(2)
    om_stats.columns = ['總需求件數', '總調貨件數', '涉及產品數量']
    om_stats['轉貨行數'] = transfer_suggestions_df.groupby('OM').size()
    
    # Transfer type distribution (updated for all three modes)
    transfer_type_stats = transfer_suggestions_df.groupby('Transfer Type').agg({
        'Transfer Qty': ['sum', 'count']
    }).round(2)
    transfer_type_stats.columns = ['總件數', '涉及行數']
    
    # Receive statistics (updated with new requirements)
    receive_stats = transfer_suggestions_df.groupby('Receive Site').agg({
        'Transfer Qty': 'sum',
        'Receive Site Target Qty': 'first'
    }).round(2)
    receive_stats.columns = ['實際接收數量', '目標需求數量']
    receive_stats['需求滿足率'] = (receive_stats['實際接收數量'] / receive_stats['目標需求數量'] * 100).round(2)
    
    return {
        'total_transfer_qty': total_transfer_qty,
        'total_transfer_lines': total_transfer_lines,
        'unique_articles': unique_articles,
        'unique_oms': unique_oms,
        'article_stats': article_stats,
        'om_stats': om_stats,
        'transfer_type_stats': transfer_type_stats,
        'receive_stats': receive_stats,
        'constraint_violations': constraint_violations,
        'violation_details': violation_details
    }

def create_visualization(stats, transfer_suggestions_df, mode):
    """Create matplotlib visualization based on mode"""
    if transfer_suggestions_df.empty:
        return None
    
    fig, ax = plt.subplots(figsize=(12, 8))
    
    # Prepare data for visualization
    om_transfer_data = transfer_suggestions_df.groupby(['OM', 'Transfer Type'])['Transfer Qty'].sum().unstack(fill_value=0)
    
    # Add receive data
    receive_data = transfer_suggestions_df.groupby('Receive Site')['Transfer Qty'].sum()
    receive_by_om = transfer_suggestions_df.drop_duplicates('Receive Site').set_index('Receive Site')['OM']
    receive_by_om = receive_by_om[receive_data.index]
    receive_by_type = pd.DataFrame({'Actual Receive Qty': receive_data.values}, index=receive_by_om.values)
    receive_by_type = receive_by_type.groupby(level=0).sum()
    
    # Add target data
    target_data = transfer_suggestions_df.drop_duplicates('Receive Site').groupby('OM')['Receive Site Target Qty'].sum()
    
    # Combine all data
    combined_data = om_transfer_data.join(receive_by_type).join(target_data).fillna(0)
    
    # Create bar chart
    if mode == 'conservative':
        # Conservative mode: 4 bars
        combined_data.plot(kind='bar', ax=ax, width=0.8)
    elif mode == 'enhanced':
        # Enhanced mode: 5 bars
        combined_data.plot(kind='bar', ax=ax, width=0.8)
    else:  # special mode
        # Special Enhanced mode: 5 bars (different types)
        combined_data.plot(kind='bar', ax=ax, width=0.8)
    
    ax.set_title('Transfer Receive Analysis', fontsize=16, fontweight='bold')
    ax.set_xlabel('OM Unit', fontsize=12)
    ax.set_ylabel('Transfer Quantity', fontsize=12)
    ax.legend(title='Transfer Type', bbox_to_anchor=(1.05, 1), loc='upper left')
    ax.grid(axis='y', alpha=0.3)
    
    plt.xticks(rotation=45)
    plt.tight_layout()
    
    return fig

def export_to_excel(transfer_suggestions_df, stats):
    """Export results to Excel file"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Worksheet 1: Transfer Suggestions
        transfer_suggestions_df.to_excel(writer, sheet_name='轉貨建議', index=False)
        
        # Worksheet 2: Statistics Summary
        workbook = writer.book
        stats_worksheet = workbook.add_worksheet('統計摘要')
        
        # KPI Overview
        stats_worksheet.write('A1', 'KPI概覽')
        stats_worksheet.write('A2', '總轉貨建議數量')
        stats_worksheet.write('B2', stats['total_transfer_qty'])
        stats_worksheet.write('A3', '總轉貨件數')
        stats_worksheet.write('B3', stats['total_transfer_lines'])
        stats_worksheet.write('A4', '涉及產品數量')
        stats_worksheet.write('B4', stats['unique_articles'])
        stats_worksheet.write('A5', '涉及OM數量')
        stats_worksheet.write('B5', stats['unique_oms'])
        
        # Article Statistics
        row = 8
        stats_worksheet.write(f'A{row}', '按產品統計')
        row += 1
        if not stats['article_stats'].empty:
            stats['article_stats'].to_excel(writer, sheet_name='統計摘要', startrow=row, startcol=0)
            row += len(stats['article_stats']) + 4
        
        # OM Statistics
        stats_worksheet.write(f'A{row}', '按OM統計')
        row += 1
        if not stats['om_stats'].empty:
            stats['om_stats'].to_excel(writer, sheet_name='統計摘要', startrow=row, startcol=0)
            row += len(stats['om_stats']) + 4
        
        # Transfer Type Distribution
        stats_worksheet.write(f'A{row}', '轉出類型分佈')
        row += 1
        if not stats['transfer_type_stats'].empty:
            stats['transfer_type_stats'].to_excel(writer, sheet_name='統計摘要', startrow=row, startcol=0)
            row += len(stats['transfer_type_stats']) + 4
        
        # Receive Type Distribution
        stats_worksheet.write(f'A{row}', '接收類型分佈')
        row += 1
        if not stats['receive_stats'].empty:
            stats['receive_stats'].to_excel(writer, sheet_name='統計摘要', startrow=row, startcol=0)
    
    output.seek(0)
    return output

# Main application
def main():
    # Page header
    st.title("📦 強制指定店鋪轉貨系統")
    st.markdown("---")
    
    # Sidebar
    st.sidebar.header("系統資訊")
    st.sidebar.info("""
    **版本：v1.0**
    **開發者:Ricky**
    
    **核心功能：**  
    - ✅ ND/RF類型智慧識別
    - ✅ 優先順序轉貨
    - ✅ 統計分析和圖表
    - ✅ Excel格式匯出
    """)
    
    # File upload section
    st.header("1. 資料上傳")
    uploaded_file = st.file_uploader(
        "請上傳Excel檔案 (.xlsx, .xls)",
        type=['xlsx', 'xls'],
        help="檔案必須包含所有必要的欄位：Article, Article Description, RP Type, Site, OM, MOQ, SaSa Net Stock, Target, Pending Received, Safety Stock, Last Month Sold Qty, MTD Sold Qty"
    )
    
    # Global variable to store processed data
    global df
    
    if uploaded_file is not None:
        try:
            # Read the uploaded file
            df = pd.read_excel(uploaded_file)
            
            # Validate file structure
            is_valid, message = validate_file_structure(df)
            
            if not is_valid:
                st.error(f"❌ {message}")
                return
            
            # Preprocess data
            with st.spinner("正在處理資料..."):
                df = preprocess_data(df)
            
            st.success("✅ 檔案上傳並處理成功！")
            
            # Data preview section
            st.header("2. 資料預覽")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("總記錄數", len(df))
            with col2:
                st.metric("產品數量", df['Article'].nunique())
            with col3:
                st.metric("店鋪數量", df['Site'].nunique())
            
            # Show sample data
            st.subheader("資料樣本")
            st.dataframe(df.head(10))
            
            # Transfer mode selection
            st.header("3. 轉貨模式選擇")
            mode = st.radio(
                "請選擇轉貨模式：",
                options=["conservative", "enhanced", "special"],
                format_func=lambda x: "A: 保守轉貨" if x == "conservative" else ("B: 加強轉貨" if x == "enhanced" else "C: 特強轉貨"),
                help="保守轉貨：RF類型轉出限制為50% | 加強轉貨：RF類型轉出限制為80% | 特強轉貨：RF類型轉出限制為90%並保留2件庫存"
            )
            
            # Analysis button
            st.header("4. 分析執行")
            if st.button("🚀 開始分析", type="primary", use_container_width=True):
                with st.spinner("正在生成轉貨建議..."):
                    # Identify transfer candidates
                    transfer_out_df = identify_transfer_out_candidates(df, mode)
                    transfer_in_df = identify_transfer_in_candidates(df)
                    
                    # Match transfers
                    transfer_suggestions_df = match_transfers(transfer_out_df, transfer_in_df, df)
                    
                    # Calculate statistics
                    stats = calculate_statistics(transfer_suggestions_df, mode)
                
                st.success("✅ 分析完成！")
                
                # Results section
                st.header("5. 分析結果")

                # Check for constraint violations
                if stats.get('constraint_violations', 0) > 0:
                    st.error(f"⚠️ 發現 {stats['constraint_violations']} 個約束違規：總轉出數量超過總需求數量")

                    # Show violation details in expandable section
                    with st.expander("約束違規詳情"):
                        for violation in stats.get('violation_details', []):
                            st.write(f"**產品 {violation['Article']}**:")
                            st.write(f"  - 總需求: {violation['Total Demand']}")
                            st.write(f"  - 總轉出: {violation['Total Transfer']}")
                            st.write(f"  - 違規數量: {violation['Violation']}")
                else:
                    # Add constraint compliance indicator
                    if stats.get('total_transfer_qty', 0) > 0:
                        st.success("✅ 所有轉貨建議均符合需求約束")

                # KPI metrics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("總轉貨建議數量", stats['total_transfer_qty'])
                with col2:
                    st.metric("總轉貨件數", stats['total_transfer_lines'])
                with col3:
                    st.metric("涉及產品數量", stats['unique_articles'])
                with col4:
                    st.metric("涉及OM數量", stats['unique_oms'])
                
                # Transfer suggestions table
                st.subheader("轉貨建議明細")
                if not transfer_suggestions_df.empty:
                    st.dataframe(transfer_suggestions_df, use_container_width=True)
                else:
                    # Use the new error handling function
                    error_info = handle_no_transfer_candidates(transfer_out_df, transfer_in_df, mode)

                    # Display user-friendly message
                    st.warning(f"⚠️ {error_info['message']}")

                    # Show suggestions in expandable section
                    with st.expander("疑難排解建議"):
                        st.write("**建議解決方案：**")
                        for suggestion in error_info['suggestions']:
                            st.write(f"• {suggestion}")

                    # Log the diagnostic information (for developers)
                    with st.expander("技術診斷資訊"):
                        st.json(error_info['diagnostic'])
                
                # Statistics tables
                st.subheader("統計分析")

                if not transfer_suggestions_df.empty:
                    if not stats['article_stats'].empty:
                        st.write("**按產品統計**")
                        st.dataframe(stats['article_stats'])

                    if not stats['om_stats'].empty:
                        st.write("**按OM統計**")
                        st.dataframe(stats['om_stats'])

                    if not stats['transfer_type_stats'].empty:
                        st.write("**轉出類型分佈**")
                        st.dataframe(stats['transfer_type_stats'])

                    if not stats['receive_stats'].empty:
                        st.write("**接收類型結果**")
                        st.dataframe(stats['receive_stats'])

                    # Visualization
                    st.subheader("數據視覺化")
                    fig = create_visualization(stats, transfer_suggestions_df, mode)
                    if fig:
                        st.pyplot(fig)
                    else:
                        st.info("沒有足夠的數據生成圖表")
                else:
                    # Show message when no data available
                    st.info("📊 沒有轉貨建議資料，無法生成統計分析和圖表")
                
                # Export section
                st.header("6. 匯出功能")

                if not transfer_suggestions_df.empty:
                    # Generate Excel file
                    excel_data = export_to_excel(transfer_suggestions_df, stats)

                    # Create download button
                    current_date = datetime.now().strftime("%Y%m%d")
                    filename = f"強制轉貨建議_{current_date}.xlsx"

                    st.download_button(
                        label="📥 下載Excel報告",
                        data=excel_data,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info("📋 沒有轉貨建議資料，無法產生Excel報告")
        
        except Exception as e:
            st.error(f"❌ 處理檔案時發生錯誤：{str(e)}")
            st.error("請檢查檔案格式和內容是否符合要求")

if __name__ == "__main__":
    main()