import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import io
import openpyxl
from openpyxl import Workbook

# Set page configuration
st.set_page_config(
    page_title="Mandatory Shop-to-Shop Transfer System",
    page_icon="📦",
    layout="wide"
)

# Constants
REQUIRED_COLUMNS = [
    'Article', 'Article Description', 'RP Type', 'Site', 'OM', 'MOQ',
    'SaSa Net Stock', 'Target', 'Pending Received', 'Safety Stock',
    'Last Month Sold Qty', 'MTD Sold Qty'
]

# Sidebar
st.sidebar.header("System Information")
st.sidebar.info("""
**Version: v1.0**
**Developer: Ricky**

**Core Features:**
- ✅ ND/RF Type Smart Identification
- ✅ Priority Order Transfer
- ✅ Statistical Analysis and Charts
- ✅ Excel Format Export
""")

# Main title
st.title("📦 Mandatory Shop-to-Shop Transfer System")

# Initialize session state
if 'data' not in st.session_state:
    st.session_state.data = None
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'transfer_results' not in st.session_state:
    st.session_state.transfer_results = None
if 'mode' not in st.session_state:
    st.session_state.mode = 'A'

def load_data(uploaded_file):
    """Load and validate Excel data"""
    try:
        # Read Excel file
        df = pd.read_excel(uploaded_file, engine='openpyxl')

        # Check for required columns
        missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
        if missing_cols:
            st.error(f"Missing required columns: {', '.join(missing_cols)}")
            return None

        # Validate data types and convert
        df = df[REQUIRED_COLUMNS].copy()

        # Convert data types
        df['Article'] = df['Article'].astype(str)
        numeric_cols = ['MOQ', 'SaSa Net Stock', 'Target', 'Pending Received',
                       'Safety Stock', 'Last Month Sold Qty', 'MTD Sold Qty']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

        # Validate RP Type
        valid_rp_types = ['ND', 'RF']
        invalid_rp = df[~df['RP Type'].isin(valid_rp_types)]
        if not invalid_rp.empty:
            st.warning(f"Found invalid RP Type values. Valid values are ND or RF. Invalid rows: {len(invalid_rp)}")

        return df

    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return None

def preprocess_data(df):
    """Preprocess the data according to business rules"""
    df = df.copy()

    # Add Notes column for logging
    df['Notes'] = ''

    # Fix negative values
    numeric_cols = ['SaSa Net Stock', 'Pending Received', 'Safety Stock',
                   'Last Month Sold Qty', 'MTD Sold Qty']
    for col in numeric_cols:
        negative_mask = df[col] < 0
        if negative_mask.any():
            df.loc[negative_mask, col] = 0
            df.loc[negative_mask, 'Notes'] += f'{col} corrected from negative to 0; '

    # Cap extreme sales values
    sales_cols = ['Last Month Sold Qty', 'MTD Sold Qty']
    for col in sales_cols:
        extreme_mask = df[col] > 100000
        if extreme_mask.any():
            df.loc[extreme_mask, col] = 100000
            df.loc[extreme_mask, 'Notes'] += f'{col} capped at 100000; '

    # Fill string columns
    string_cols = ['Article Description', 'RP Type', 'Site', 'OM']
    for col in string_cols:
        df[col] = df[col].fillna('')

    return df

def calculate_effective_sales(row):
    """Calculate effective sales quantity"""
    if row['Last Month Sold Qty'] > 0:
        return row['Last Month Sold Qty']
    else:
        return row['MTD Sold Qty']

def get_max_sales_per_article(df, article):
    """Get maximum sales for an article across all sites"""
    article_data = df[df['Article'] == article]
    return article_data.apply(calculate_effective_sales, axis=1).max()

def generate_transfer_recommendations_conservative(df):
    """Generate transfer recommendations for Mode A: Conservative Transfer"""
    df = df.copy()
    df['Effective Sales'] = df.apply(calculate_effective_sales, axis=1)

    # Calculate max sales per article
    max_sales_dict = {}
    for article in df['Article'].unique():
        max_sales_dict[article] = get_max_sales_per_article(df, article)

    # Initialize transfer candidates
    transfer_out_candidates = []
    receive_candidates = []

    # Identify transfer out candidates (Priority 1: ND type complete transfer)
    nd_mask = (df['RP Type'] == 'ND') & (df['SaSa Net Stock'] > 0)
    for _, row in df[nd_mask].iterrows():
        transfer_out_candidates.append({
            'Article': row['Article'],
            'Site': row['Site'],
            'OM': row['OM'],
            'Transfer Qty': row['SaSa Net Stock'],
            'Transfer Type': 'ND Transfer',
            'Priority': 1
        })

    # Identify transfer out candidates (Priority 2: RF type excess transfer)
    rf_mask = (df['RP Type'] == 'RF') & \
              ((df['SaSa Net Stock'] + df['Pending Received']) > df['Safety Stock']) & \
              (df['Effective Sales'] < max_sales_dict.get(row['Article'], 0))

    # Sort by sales ascending for conservative approach
    rf_candidates = df[rf_mask].copy()
    rf_candidates['Effective Sales'] = rf_candidates.apply(calculate_effective_sales, axis=1)
    rf_candidates = rf_candidates.sort_values('Effective Sales')

    for _, row in rf_candidates.iterrows():
        available_stock = row['SaSa Net Stock'] + row['Pending Received']
        base_transfer = available_stock - row['Safety Stock']
        max_transfer = available_stock * 0.5
        transfer_qty = min(base_transfer, max_transfer)
        transfer_qty = min(transfer_qty, row['SaSa Net Stock'])  # Cannot exceed actual stock

        if transfer_qty > 0:
            transfer_out_candidates.append({
                'Article': row['Article'],
                'Site': row['Site'],
                'OM': row['OM'],
                'Transfer Qty': int(transfer_qty),
                'Transfer Type': 'RF Excess Transfer',
                'Priority': 2
            })

    # Identify receive candidates
    receive_mask = df['Target'] > 0
    for _, row in df[receive_mask].iterrows():
        receive_candidates.append({
            'Article': row['Article'],
            'Site': row['Site'],
            'OM': row['OM'],
            'Target Qty': row['Target'],
            'Priority': 1
        })

    # Sort candidates by priority
    transfer_out_candidates.sort(key=lambda x: x['Priority'])
    receive_candidates.sort(key=lambda x: x['Priority'])

    # Matching algorithm
    transfers = []
    used_stock = {}  # Track used stock per site-article

    for transfer in transfer_out_candidates:
        transfer_key = (transfer['Site'], transfer['Article'])
        if transfer_key not in used_stock:
            used_stock[transfer_key] = 0

        available_qty = transfer['Transfer Qty'] - used_stock[transfer_key]
        if available_qty <= 0:
            continue

        # Find matching receives
        for receive in receive_candidates:
            if (transfer['Article'] == receive['Article'] and
                transfer['OM'] == receive['OM'] and
                transfer['Site'] != receive['Site']):

                # Check total demand constraint
                total_demand = sum(r['Target Qty'] for r in receive_candidates
                                 if r['Article'] == transfer['Article'] and r['OM'] == transfer['OM'])
                current_allocated = sum(t['Receive Qty'] for t in transfers
                                      if t['Article'] == transfer['Article'] and t['OM'] == transfer['OM'])

                if current_allocated >= total_demand:
                    continue

                transfer_qty = min(available_qty, receive['Target Qty'])
                if transfer_qty > 0:
                    transfers.append({
                        'Article': transfer['Article'],
                        'Article Description': df[df['Article'] == transfer['Article']]['Article Description'].iloc[0],
                        'OM': transfer['OM'],
                        'Transfer Site': transfer['Site'],
                        'Transfer Qty': transfer_qty,
                        'Transfer Site Original Stock': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['SaSa Net Stock'].iloc[0],
                        'Transfer Site After Transfer Stock': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['SaSa Net Stock'].iloc[0] - transfer_qty,
                        'Transfer Site Safety Stock': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['Safety Stock'].iloc[0],
                        'Transfer Site MOQ': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['MOQ'].iloc[0],
                        'Receive Site': receive['Site'],
                        'Receive Site Target Qty': receive['Target Qty'],
                        'Transfer Type': transfer['Transfer Type'],
                        'Receive Qty': transfer_qty,
                        'Notes': ''
                    })

                    used_stock[transfer_key] += transfer_qty
                    receive['Target Qty'] -= transfer_qty
                    available_qty -= transfer_qty

                    if available_qty <= 0:
                        break

    return transfers

def generate_transfer_recommendations_enhanced(df):
    """Generate transfer recommendations for Mode B: Enhanced Transfer"""
    df = df.copy()
    df['Effective Sales'] = df.apply(calculate_effective_sales, axis=1)

    # Calculate max sales per article
    max_sales_dict = {}
    for article in df['Article'].unique():
        max_sales_dict[article] = get_max_sales_per_article(df, article)

    # Initialize transfer candidates
    transfer_out_candidates = []
    receive_candidates = []

    # Identify transfer out candidates (Priority 1: ND type complete transfer)
    nd_mask = (df['RP Type'] == 'ND') & (df['SaSa Net Stock'] > 0)
    for _, row in df[nd_mask].iterrows():
        transfer_out_candidates.append({
            'Article': row['Article'],
            'Site': row['Site'],
            'OM': row['OM'],
            'Transfer Qty': row['SaSa Net Stock'],
            'Transfer Type': 'ND Transfer',
            'Priority': 1
        })

    # Identify transfer out candidates (Priority 2: RF type enhanced transfer)
    # RF類型的轉移基於MOQ和銷售表現，轉移量計算為：min(可用庫存 - MOQ, 可用庫存 * 0.9)
    rf_mask = (df['RP Type'] == 'RF') & \
              ((df['SaSa Net Stock'] + df['Pending Received']) > df['MOQ']) & \
              (df['Effective Sales'] < df['Article'].map(max_sales_dict))

    # Sort by sales ascending (lower sales sites transfer first)
    rf_candidates = df[rf_mask].copy()
    rf_candidates['Effective Sales'] = rf_candidates.apply(calculate_effective_sales, axis=1)
    rf_candidates = rf_candidates.sort_values('Effective Sales')

    for _, row in rf_candidates.iterrows():
        available_stock = row['SaSa Net Stock'] + row['Pending Received']
        base_transfer = available_stock - row['MOQ']
        max_transfer = available_stock * 0.9
        transfer_qty = min(base_transfer, max_transfer)
        transfer_qty = min(transfer_qty, row['SaSa Net Stock'])  # Cannot exceed actual stock

        if transfer_qty > 0:
            transfer_out_candidates.append({
                'Article': row['Article'],
                'Site': row['Site'],
                'OM': row['OM'],
                'Transfer Qty': int(transfer_qty),
                'Transfer Type': 'RF Enhanced Transfer',
                'Priority': 2
            })

    # Identify receive candidates
    receive_mask = df['Target'] > 0
    for _, row in df[receive_mask].iterrows():
        receive_candidates.append({
            'Article': row['Article'],
            'Site': row['Site'],
            'OM': row['OM'],
            'Target Qty': row['Target'],
            'Priority': 1
        })

    # Sort candidates by priority
    transfer_out_candidates.sort(key=lambda x: x['Priority'])
    receive_candidates.sort(key=lambda x: x['Priority'])

    # Matching algorithm (same as conservative)
    transfers = []
    used_stock = {}

    for transfer in transfer_out_candidates:
        transfer_key = (transfer['Site'], transfer['Article'])
        if transfer_key not in used_stock:
            used_stock[transfer_key] = 0

        available_qty = transfer['Transfer Qty'] - used_stock[transfer_key]
        if available_qty <= 0:
            continue

        for receive in receive_candidates:
            if (transfer['Article'] == receive['Article'] and
                transfer['OM'] == receive['OM'] and
                transfer['Site'] != receive['Site']):

                total_demand = sum(r['Target Qty'] for r in receive_candidates
                                 if r['Article'] == transfer['Article'] and r['OM'] == transfer['OM'])
                current_allocated = sum(t['Receive Qty'] for t in transfers
                                      if t['Article'] == transfer['Article'] and t['OM'] == transfer['OM'])

                if current_allocated >= total_demand:
                    continue

                transfer_qty = min(available_qty, receive['Target Qty'])
                if transfer_qty > 0:
                    transfers.append({
                        'Article': transfer['Article'],
                        'Article Description': df[df['Article'] == transfer['Article']]['Article Description'].iloc[0],
                        'OM': transfer['OM'],
                        'Transfer Site': transfer['Site'],
                        'Transfer Qty': transfer_qty,
                        'Transfer Site Original Stock': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['SaSa Net Stock'].iloc[0],
                        'Transfer Site After Transfer Stock': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['SaSa Net Stock'].iloc[0] - transfer_qty,
                        'Transfer Site Safety Stock': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['Safety Stock'].iloc[0],
                        'Transfer Site MOQ': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['MOQ'].iloc[0],
                        'Receive Site': receive['Site'],
                        'Receive Site Target Qty': receive['Target Qty'],
                        'Transfer Type': transfer['Transfer Type'],
                        'Receive Qty': transfer_qty,
                        'Notes': ''
                    })

                    used_stock[transfer_key] += transfer_qty
                    receive['Target Qty'] -= transfer_qty
                    available_qty -= transfer_qty

                    if available_qty <= 0:
                        break

    return transfers

def generate_transfer_recommendations_super(df):
    """Generate transfer recommendations for Mode C: Super Enhanced Transfer"""
    df = df.copy()
    df['Effective Sales'] = df.apply(calculate_effective_sales, axis=1)

    # Calculate max sales per article
    max_sales_dict = {}
    for article in df['Article'].unique():
        max_sales_dict[article] = get_max_sales_per_article(df, article)

    # Initialize transfer candidates
    transfer_out_candidates = []
    receive_candidates = []

    # Identify transfer out candidates (Priority 1: ND type complete transfer)
    nd_mask = (df['RP Type'] == 'ND') & (df['SaSa Net Stock'] > 0)
    for _, row in df[nd_mask].iterrows():
        transfer_out_candidates.append({
            'Article': row['Article'],
            'Site': row['Site'],
            'OM': row['OM'],
            'Transfer Qty': row['SaSa Net Stock'],
            'Transfer Type': 'ND Transfer',
            'Priority': 1
        })

    # Identify transfer out candidates (Priority 2: RF type super enhanced transfer)
    # RF類型的轉移可忽視最小庫存要求，參考銷售表現，過去銷售最多的店舖排最後出貨
    # 最大轉移量可用庫存的100%，以滿足目標需求
    rf_mask = (df['RP Type'] == 'RF') & (df['SaSa Net Stock'] > 0)

    # Sort by sales ascending (lower sales sites transfer first, highest sales last)
    rf_candidates = df[rf_mask].copy()
    rf_candidates['Effective Sales'] = rf_candidates.apply(calculate_effective_sales, axis=1)
    rf_candidates = rf_candidates.sort_values('Effective Sales', ascending=True)

    for _, row in rf_candidates.iterrows():
        # 可轉移全部實際庫存，不需保留任何庫存
        transfer_qty = max(0, row['SaSa Net Stock'])

        if transfer_qty > 0:
            transfer_out_candidates.append({
                'Article': row['Article'],
                'Site': row['Site'],
                'OM': row['OM'],
                'Transfer Qty': int(transfer_qty),
                'Transfer Type': 'RF Super Enhanced Transfer',
                'Priority': 2,
                'Effective Sales': row['Effective Sales']  # 記錄銷售量用於排序
            })

    # Identify receive candidates
    receive_mask = df['Target'] > 0
    for _, row in df[receive_mask].iterrows():
        receive_candidates.append({
            'Article': row['Article'],
            'Site': row['Site'],
            'OM': row['OM'],
            'Target Qty': row['Target'],
            'Priority': 1
        })

    # Sort candidates by priority
    transfer_out_candidates.sort(key=lambda x: x['Priority'])
    receive_candidates.sort(key=lambda x: x['Priority'])

    # Matching algorithm (same as before)
    transfers = []
    used_stock = {}

    for transfer in transfer_out_candidates:
        transfer_key = (transfer['Site'], transfer['Article'])
        if transfer_key not in used_stock:
            used_stock[transfer_key] = 0

        available_qty = transfer['Transfer Qty'] - used_stock[transfer_key]
        if available_qty <= 0:
            continue

        for receive in receive_candidates:
            if (transfer['Article'] == receive['Article'] and
                transfer['OM'] == receive['OM'] and
                transfer['Site'] != receive['Site']):

                total_demand = sum(r['Target Qty'] for r in receive_candidates
                                 if r['Article'] == transfer['Article'] and r['OM'] == transfer['OM'])
                current_allocated = sum(t['Receive Qty'] for t in transfers
                                      if t['Article'] == transfer['Article'] and t['OM'] == transfer['OM'])

                if current_allocated >= total_demand:
                    continue

                transfer_qty = min(available_qty, receive['Target Qty'])
                if transfer_qty > 0:
                    transfers.append({
                        'Article': transfer['Article'],
                        'Article Description': df[df['Article'] == transfer['Article']]['Article Description'].iloc[0],
                        'OM': transfer['OM'],
                        'Transfer Site': transfer['Site'],
                        'Transfer Qty': transfer_qty,
                        'Transfer Site Original Stock': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['SaSa Net Stock'].iloc[0],
                        'Transfer Site After Transfer Stock': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['SaSa Net Stock'].iloc[0] - transfer_qty,
                        'Transfer Site Safety Stock': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['Safety Stock'].iloc[0],
                        'Transfer Site MOQ': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['MOQ'].iloc[0],
                        'Receive Site': receive['Site'],
                        'Receive Site Target Qty': receive['Target Qty'],
                        'Transfer Type': transfer['Transfer Type'],
                        'Receive Qty': transfer_qty,
                        'Notes': ''
                    })

                    used_stock[transfer_key] += transfer_qty
                    receive['Target Qty'] -= transfer_qty
                    available_qty -= transfer_qty

                    if available_qty <= 0:
                        break

    return transfers

def calculate_statistics(transfers, df):
    """Calculate comprehensive statistics"""
    stats = {}

    # Basic KPIs
    stats['total_recommendations'] = len(transfers)
    stats['total_transfer_qty'] = sum(t['Transfer Qty'] for t in transfers)
    stats['unique_articles'] = len(set(t['Article'] for t in transfers))
    stats['unique_oms'] = len(set(t['OM'] for t in transfers))

    # By Article statistics
    article_stats = []
    for article in set(t['Article'] for t in transfers):
        article_transfers = [t for t in transfers if t['Article'] == article]
        total_demand = sum(df[(df['Article'] == article) & (df['Target'] > 0)]['Target'])
        total_transfer = sum(t['Transfer Qty'] for t in article_transfers)
        fulfillment_rate = (total_transfer / total_demand * 100) if total_demand > 0 else 0

        article_stats.append({
            'Article': article,
            'Total Demand Qty': total_demand,
            'Total Transfer Qty': total_transfer,
            'Transfer Lines': len(article_transfers),
            'Fulfillment Rate (%)': round(fulfillment_rate, 2)
        })

    # By OM statistics
    om_stats = []
    for om in set(t['OM'] for t in transfers):
        om_transfers = [t for t in transfers if t['OM'] == om]
        total_demand = sum(df[(df['OM'] == om) & (df['Target'] > 0)]['Target'])
        total_transfer = sum(t['Transfer Qty'] for t in om_transfers)
        unique_articles = len(set(t['Article'] for t in om_transfers))

        om_stats.append({
            'OM': om,
            'Total Transfer Qty': total_transfer,
            'Total Demand Qty': total_demand,
            'Transfer Lines': len(om_transfers),
            'Unique Articles': unique_articles
        })

    # Transfer type distribution
    transfer_types = {}
    for transfer in transfers:
        ttype = transfer['Transfer Type']
        if ttype not in transfer_types:
            transfer_types[ttype] = {'qty': 0, 'lines': 0}
        transfer_types[ttype]['qty'] += transfer['Transfer Qty']
        transfer_types[ttype]['lines'] += 1

    # Receive statistics
    receive_stats = []
    for site in set(t['Receive Site'] for t in transfers):
        site_transfers = [t for t in transfers if t['Receive Site'] == site]
        total_target = df[df['Site'] == site]['Target'].sum()
        total_received = sum(t['Receive Qty'] for t in site_transfers)

        receive_stats.append({
            'Site': site,
            'Total Target Qty': total_target,
            'Total Received Qty': total_received
        })

    return {
        'basic': stats,
        'by_article': article_stats,
        'by_om': om_stats,
        'transfer_types': transfer_types,
        'receive_stats': receive_stats
    }

def create_visualization(transfers, mode, df):
    """Create matplotlib visualization based on mode"""
    if not transfers:
        return None

    # Prepare data
    om_data = {}
    for transfer in transfers:
        om = transfer['OM']
        if om not in om_data:
            om_data[om] = {
                'ND Transfer': 0,
                'RF Transfer': 0,
                'Demand': 0,
                'Received': 0
            }

        if 'ND' in transfer['Transfer Type']:
            om_data[om]['ND Transfer'] += transfer['Transfer Qty']
        else:
            om_data[om]['RF Transfer'] += transfer['Transfer Qty']

        om_data[om]['Received'] += transfer['Receive Qty']

    # Add demand data - total demand for the OM
    for om in om_data:
        om_data[om]['Demand'] = df[(df['OM'] == om) & (df['Target'] > 0)]['Target'].sum()

    # Create plot
    fig, ax = plt.subplots(figsize=(12, 6))

    oms = list(om_data.keys())
    nd_transfer = [om_data[om]['ND Transfer'] for om in oms]
    rf_transfer = [om_data[om]['RF Transfer'] for om in oms]
    demand = [om_data[om]['Demand'] for om in oms]
    received = [om_data[om]['Received'] for om in oms]

    x = np.arange(len(oms))
    width = 0.2

    if mode == 'A':
        ax.bar(x - width*1.5, nd_transfer, width, label='ND Transfer Qty', color='blue')
        ax.bar(x - width/2, rf_transfer, width, label='RF Excess Transfer Qty', color='green')
        ax.bar(x + width/2, demand, width, label='Demand Qty', color='red')
        ax.bar(x + width*1.5, received, width, label='Actual Received Qty', color='orange')
    elif mode == 'B':
        ax.bar(x - width*2, nd_transfer, width, label='ND Transfer Qty', color='blue')
        ax.bar(x - width, rf_transfer, width, label='RF Enhanced Transfer Qty', color='green')
        ax.bar(x, demand, width, label='Demand Qty', color='red')
        ax.bar(x + width, received, width, label='Actual Received Qty', color='orange')
    else:  # Mode C
        ax.bar(x - width*2, nd_transfer, width, label='ND Transfer Qty', color='blue')
        ax.bar(x - width, rf_transfer, width, label='RF Super Enhanced Transfer Qty', color='green')
        ax.bar(x, demand, width, label='Demand Qty', color='red')
        ax.bar(x + width, received, width, label='Actual Received Qty', color='orange')

    ax.set_xlabel('OM Units')
    ax.set_ylabel('Transfer Quantity')
    ax.set_title('Transfer Receive Analysis')
    ax.set_xticks(x)
    ax.set_xticklabels(oms)
    ax.legend()

    plt.tight_layout()
    return fig

def export_to_excel(transfers, stats):
    """Export results to Excel with two sheets"""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Transfer Recommendations
        if transfers:
            transfer_df = pd.DataFrame(transfers)
            # Reorder columns as specified
            columns_order = [
                'Article', 'Article Description', 'OM', 'Transfer Site', 'Transfer Qty',
                'Transfer Site Original Stock', 'Transfer Site After Transfer Stock',
                'Transfer Site Safety Stock', 'Transfer Site MOQ', 'Receive Site',
                'Receive Site Target Qty', 'Notes'
            ]
            transfer_df = transfer_df[columns_order]
            transfer_df.to_excel(writer, sheet_name='Transfer Recommendations', index=False)

        # Sheet 2: Statistics Summary
        # Basic KPIs
        basic_stats = pd.DataFrame([{
            'Metric': 'Total Recommendations',
            'Value': stats['basic']['total_recommendations']
        }, {
            'Metric': 'Total Transfer Quantity',
            'Value': stats['basic']['total_transfer_qty']
        }, {
            'Metric': 'Unique Articles',
            'Value': stats['basic']['unique_articles']
        }, {
            'Metric': 'Unique OMs',
            'Value': stats['basic']['unique_oms']
        }])

        start_row = 0
        basic_stats.to_excel(writer, sheet_name='Statistics Summary', startrow=start_row, index=False)

        # By Article
        start_row += len(basic_stats) + 3
        if stats['by_article']:
            article_df = pd.DataFrame(stats['by_article'])
            article_df.to_excel(writer, sheet_name='Statistics Summary', startrow=start_row, index=False)

        # By OM
        start_row += len(stats['by_article']) + 3
        if stats['by_om']:
            om_df = pd.DataFrame(stats['by_om'])
            om_df.to_excel(writer, sheet_name='Statistics Summary', startrow=start_row, index=False)

        # Transfer Types
        start_row += len(stats['by_om']) + 3
        if stats['transfer_types']:
            type_data = []
            for ttype, data in stats['transfer_types'].items():
                type_data.append({
                    'Transfer Type': ttype,
                    'Total Quantity': data['qty'],
                    'Total Lines': data['lines']
                })
            type_df = pd.DataFrame(type_data)
            type_df.to_excel(writer, sheet_name='Statistics Summary', startrow=start_row, index=False)

        # Receive Stats
        start_row += len(type_data) + 3
        if stats['receive_stats']:
            receive_df = pd.DataFrame(stats['receive_stats'])
            receive_df.to_excel(writer, sheet_name='Statistics Summary', startrow=start_row, index=False)

    output.seek(0)
    return output

# Main UI
st.header("1. Data Upload")
uploaded_file = st.file_uploader("Upload Excel file", type=['xlsx', 'xls'])

if uploaded_file is not None:
    with st.spinner("Loading and validating data..."):
        data = load_data(uploaded_file)
        if data is not None:
            st.session_state.data = data
            st.success(f"Data loaded successfully! {len(data)} rows processed.")

            # Data preview
            st.header("2. Data Preview")
            st.subheader("Sample Data")
            st.dataframe(data.head(10))

            st.subheader("Basic Statistics")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Rows", len(data))
            with col2:
                st.metric("Unique Articles", data['Article'].nunique())
            with col3:
                st.metric("Unique Sites", data['Site'].nunique())
            with col4:
                st.metric("Unique OMs", data['OM'].nunique())

            # Preprocess data
            with st.spinner("Preprocessing data..."):
                processed_data = preprocess_data(data)
                st.session_state.processed_data = processed_data

            # Mode selection
            st.header("3. Transfer Mode Selection")
            mode = st.radio(
                "Select Transfer Mode:",
                ['A: Conservative Transfer', 'B: Enhanced Transfer', 'C: Super Enhanced Transfer'],
                index=0
            )
            st.session_state.mode = mode[0]

            # Generate recommendations
            if st.button("Generate Transfer Recommendations", type="primary"):
                with st.spinner("Generating recommendations..."):
                    if mode.startswith('A'):
                        transfers = generate_transfer_recommendations_conservative(processed_data)
                    elif mode.startswith('B'):
                        transfers = generate_transfer_recommendations_enhanced(processed_data)
                    else:  # Mode C
                        transfers = generate_transfer_recommendations_super(processed_data)

                    st.session_state.transfer_results = transfers

                    if transfers:
                        st.success(f"Generated {len(transfers)} transfer recommendations!")

                        # Statistics
                        st.header("4. Analysis Results")
                        stats = calculate_statistics(transfers, processed_data)

                        # KPI Cards
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("Total Recommendations", stats['basic']['total_recommendations'])
                        with col2:
                            st.metric("Total Transfer Qty", stats['basic']['total_transfer_qty'])
                        with col3:
                            st.metric("Unique Articles", stats['basic']['unique_articles'])
                        with col4:
                            st.metric("Unique OMs", stats['basic']['unique_oms'])

                        # Transfer Results Table
                        st.subheader("Transfer Recommendations")
                        transfer_df = pd.DataFrame(transfers)
                        st.dataframe(transfer_df)

                        # Statistics Tables
                        st.subheader("Statistics by Article")
                        if stats['by_article']:
                            st.dataframe(pd.DataFrame(stats['by_article']))

                        st.subheader("Statistics by OM")
                        if stats['by_om']:
                            st.dataframe(pd.DataFrame(stats['by_om']))

                        st.subheader("Transfer Type Distribution")
                        if stats['transfer_types']:
                            type_data = []
                            for ttype, data in stats['transfer_types'].items():
                                type_data.append({
                                    'Type': ttype,
                                    'Total Qty': data['qty'],
                                    'Lines': data['lines']
                                })
                            st.dataframe(pd.DataFrame(type_data))

                        # Visualization
                        st.subheader("Transfer Analysis Chart")
                        fig = create_visualization(transfers, st.session_state.mode, processed_data)
                        if fig:
                            st.pyplot(fig)

                        # Export
                        st.header("5. Export Results")
                        excel_data = export_to_excel(transfers, stats)
                        st.download_button(
                            label="📥 Download Excel File",
                            data=excel_data.getvalue(),
                            file_name=f"Mandatory_Transfer_Recommendations_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="excel_download"
                        )
                    else:
                        st.warning("No transfer recommendations generated. Please check your data and try different mode.")

else:
    st.info("Please upload an Excel file to get started.")

# Footer
st.markdown("---")
st.markdown("*Developed by Ricky - Version 1.0*")