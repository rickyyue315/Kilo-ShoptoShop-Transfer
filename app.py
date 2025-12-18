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
    page_title="店舖間強制轉移系統",
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
st.sidebar.header("系統資訊")
st.sidebar.info("""
**版本: v2.0**
**開發者: Ricky**

**核心功能:**
- ✅ ND/RF 類型智能識別
- ✅ 優先級訂單轉移
- ✅ 統計分析與圖表
- ✅ Excel 格式匯出
""")

# Main title
st.title("📦 店舖間強制轉移系統")

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
    """載入並驗證Excel資料"""
    try:
        # 讀取Excel檔案
        df = pd.read_excel(uploaded_file, engine='openpyxl')

        # 檢查必要欄位
        missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
        if missing_cols:
            st.error(f"缺少必要欄位: {', '.join(missing_cols)}")
            return None

        # 驗證資料類型並轉換
        df = df[REQUIRED_COLUMNS].copy()

        # 轉換資料類型
        df['Article'] = df['Article'].astype(str)
        numeric_cols = ['MOQ', 'SaSa Net Stock', 'Target', 'Pending Received',
                       'Safety Stock', 'Last Month Sold Qty', 'MTD Sold Qty']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

        # 驗證RP Type
        valid_rp_types = ['ND', 'RF']
        invalid_rp = df[~df['RP Type'].isin(valid_rp_types)]
        if not invalid_rp.empty:
            st.warning(f"發現無效的RP Type值。有效值為ND或RF。無效行數: {len(invalid_rp)}")

        return df

    except Exception as e:
        st.error(f"載入檔案時發生錯誤: {str(e)}")
        return None

def preprocess_data(df):
    """根據業務規則預處理資料"""
    df = df.copy()

    # 新增註記欄位
    df['Notes'] = ''

    # 修正負值
    numeric_cols = ['SaSa Net Stock', 'Pending Received', 'Safety Stock',
                   'Last Month Sold Qty', 'MTD Sold Qty']
    for col in numeric_cols:
        negative_mask = df[col] < 0
        if negative_mask.any():
            df.loc[negative_mask, col] = 0
            df.loc[negative_mask, 'Notes'] += f'{col} 從負值修正為0; '

    # 限制極端銷售值
    sales_cols = ['Last Month Sold Qty', 'MTD Sold Qty']
    for col in sales_cols:
        extreme_mask = df[col] > 100000
        if extreme_mask.any():
            df.loc[extreme_mask, col] = 100000
            df.loc[extreme_mask, 'Notes'] += f'{col} 限制為100000; '

    # 填充字串欄位
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
              (df['Effective Sales'] < df['Article'].map(max_sales_dict))

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
                        'Transfer Site RP Type': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['RP Type'].iloc[0],
                        'Transfer Site Last Month Sold Qty': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['Last Month Sold Qty'].iloc[0],
                        'Transfer Site MTD Sold Qty': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['MTD Sold Qty'].iloc[0],
                        'Receive Site': receive['Site'],
                        'Receive Site Target Qty': receive['Target Qty'],
                        'Receive Site RP Type': df[(df['Site'] == receive['Site']) & (df['Article'] == receive['Article'])]['RP Type'].iloc[0],
                        'Receive Site Last Month Sold Qty': df[(df['Site'] == receive['Site']) & (df['Article'] == receive['Article'])]['Last Month Sold Qty'].iloc[0],
                        'Receive Site MTD Sold Qty': df[(df['Site'] == receive['Site']) & (df['Article'] == receive['Article'])]['MTD Sold Qty'].iloc[0],
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
                        'Transfer Site RP Type': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['RP Type'].iloc[0],
                        'Transfer Site Last Month Sold Qty': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['Last Month Sold Qty'].iloc[0],
                        'Transfer Site MTD Sold Qty': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['MTD Sold Qty'].iloc[0],
                        'Receive Site': receive['Site'],
                        'Receive Site Target Qty': receive['Target Qty'],
                        'Receive Site RP Type': df[(df['Site'] == receive['Site']) & (df['Article'] == receive['Article'])]['RP Type'].iloc[0],
                        'Receive Site Last Month Sold Qty': df[(df['Site'] == receive['Site']) & (df['Article'] == receive['Article'])]['Last Month Sold Qty'].iloc[0],
                        'Receive Site MTD Sold Qty': df[(df['Site'] == receive['Site']) & (df['Article'] == receive['Article'])]['MTD Sold Qty'].iloc[0],
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

    # Matching algorithm for Mode C - 允許不同OM組別調撥，只限制HD不能去HA,HB,HC組別
    transfers = []
    used_stock = {}
    
    # 計算每個商品的總需求（跨所有OM組別）
    article_total_demand = {}
    for article in df['Article'].unique():
        article_total_demand[article] = df[(df['Article'] == article) & (df['Target'] > 0)]['Target'].sum()

    for transfer in transfer_out_candidates:
        transfer_key = (transfer['Site'], transfer['Article'])
        if transfer_key not in used_stock:
            used_stock[transfer_key] = 0

        available_qty = transfer['Transfer Qty'] - used_stock[transfer_key]
        if available_qty <= 0:
            continue

        for receive in receive_candidates:
            # Mode C: 只限制HD不能去HA,HB,HC組別
            transfer_om = transfer['OM']
            receive_om = receive['OM']
            
            # 檢查限制條件：如果轉出店是HD，接收店不能是HA,HB,HC
            if transfer_om == 'HD' and receive_om in ['HA', 'HB', 'HC']:
                continue
                
            # 檢查是否同一店舖
            if transfer['Site'] == receive['Site']:
                continue

            # 檢查商品是否相同
            if transfer['Article'] != receive['Article']:
                continue

            # 檢查總需求限制（所有接收店的總需求）
            total_demand = article_total_demand.get(transfer['Article'], 0)
            current_allocated = sum(t['Receive Qty'] for t in transfers
                                  if t['Article'] == transfer['Article'])

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
                    'Transfer Site RP Type': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['RP Type'].iloc[0],
                    'Transfer Site Last Month Sold Qty': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['Last Month Sold Qty'].iloc[0],
                    'Transfer Site MTD Sold Qty': df[(df['Site'] == transfer['Site']) & (df['Article'] == transfer['Article'])]['MTD Sold Qty'].iloc[0],
                    'Receive Site': receive['Site'],
                    'Receive Site Target Qty': receive['Target Qty'],
                    'Receive Site RP Type': df[(df['Site'] == receive['Site']) & (df['Article'] == receive['Article'])]['RP Type'].iloc[0],
                    'Receive Site Last Month Sold Qty': df[(df['Site'] == receive['Site']) & (df['Article'] == receive['Article'])]['Last Month Sold Qty'].iloc[0],
                    'Receive Site MTD Sold Qty': df[(df['Site'] == receive['Site']) & (df['Article'] == receive['Article'])]['MTD Sold Qty'].iloc[0],
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
    """根據模式建立matplotlib視覺化圖表"""
    if not transfers:
        return None

    # 準備資料
    om_data = {}
    for transfer in transfers:
        om = transfer['OM']
        if om not in om_data:
            om_data[om] = {
                'ND Transfer': 0,
                'RF Transfer': 0,
                'Demand': 0,
                'Actual Received': 0
            }

        if 'ND' in transfer['Transfer Type']:
            om_data[om]['ND Transfer'] += transfer['Transfer Qty']
        else:
            om_data[om]['RF Transfer'] += transfer['Transfer Qty']

        om_data[om]['Actual Received'] += transfer['Receive Qty']

    # 新增需求資料 - 該OM的總需求
    for om in om_data:
        om_data[om]['Demand'] = df[(df['OM'] == om) & (df['Target'] > 0)]['Target'].sum()

    # 建立圖表
    fig, ax = plt.subplots(figsize=(12, 6))

    oms = list(om_data.keys())
    nd_transfer = [om_data[om]['ND Transfer'] for om in oms]
    rf_transfer = [om_data[om]['RF Transfer'] for om in oms]
    demand = [om_data[om]['Demand'] for om in oms]
    received = [om_data[om]['Actual Received'] for om in oms]

    x = np.arange(len(oms))
    width = 0.2

    if mode == 'A':
        ax.bar(x - width*1.5, nd_transfer, width, label='ND Transfer', color='blue')
        ax.bar(x - width/2, rf_transfer, width, label='RF Excess Transfer', color='green')
        ax.bar(x + width/2, demand, width, label='Demand', color='red')
        ax.bar(x + width*1.5, received, width, label='Actual Received', color='orange')
    elif mode == 'B':
        ax.bar(x - width*2, nd_transfer, width, label='ND Transfer', color='blue')
        ax.bar(x - width, rf_transfer, width, label='RF Enhanced Transfer', color='green')
        ax.bar(x, demand, width, label='Demand', color='red')
        ax.bar(x + width, received, width, label='Actual Received', color='orange')
    else:  # Mode C
        ax.bar(x - width*2, nd_transfer, width, label='ND Transfer', color='blue')
        ax.bar(x - width, rf_transfer, width, label='RF Super Enhanced Transfer', color='green')
        ax.bar(x, demand, width, label='Demand', color='red')
        ax.bar(x + width, received, width, label='Actual Received', color='orange')

    ax.set_xlabel('OM Group')
    ax.set_ylabel('Transfer Quantity')
    ax.set_title('Transfer Analysis Chart')
    ax.set_xticks(x)
    ax.set_xticklabels(oms)
    ax.legend()

    plt.tight_layout()
    return fig

def export_to_excel(transfers, stats, df):
    """將結果匯出到Excel，包含兩個工作表"""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 工作表1: 轉移建議
        if transfers:
            # 準備轉移數據 - 根據參考格式調整
            transfer_data = []
            for transfer in transfers:
                # 獲取接收店舖的原始庫存
                receive_original_stock = df[(df['Site'] == transfer['Receive Site']) &
                                           (df['Article'] == transfer['Article'])]['SaSa Net Stock'].iloc[0] if len(df[(df['Site'] == transfer['Receive Site']) & (df['Article'] == transfer['Article'])]) > 0 else 0
                
                # 生成Remark和Notes內容
                transfer_type = transfer['Transfer Type']
                if 'ND' in transfer_type:
                    remark = "ND轉出 → 緊急缺貨補貨" if transfer['Receive Qty'] > 0 else "ND轉出"
                    notes = f"【轉出分類: {transfer_type}】 | 【接收分類: 緊急缺貨補貨】 | 【轉出優先級: ND轉出】 | 【接收優先級: 接收(最高優先級)】"
                else:
                    remark = f"{transfer_type} → 潛在缺貨補貨"
                    notes = f"【轉出分類: {transfer_type}】 | 【接收分類: 潛在缺貨補貨】 | 【轉出優先級: RF轉出】 | 【接收優先級: 接收(一般優先級)】"
                
                row = {
                    'Article': transfer['Article'],
                    'Product Desc': transfer['Article Description'],
                    'Transfer OM': transfer['OM'],
                    'Transfer Site': transfer['Transfer Site'],
                    'Receive OM': transfer['OM'],  # Mode A/B為相同OM，C模式可能不同
                    'Receive Site': transfer['Receive Site'],
                    'Transfer Qty': transfer['Transfer Qty'],
                    'Transfer Original Stock': transfer['Transfer Site Original Stock'],
                    'Transfer After Transfer Stock': transfer['Transfer Site After Transfer Stock'],
                    'Transfer Safety Stock': transfer['Transfer Site Safety Stock'],
                    'Transfer MOQ': transfer['Transfer Site MOQ'],
                    'Remark': remark,
                    'Notes': notes,
                    'Transfer Site Last Month Sold Qty': transfer.get('Transfer Site Last Month Sold Qty', 0),
                    'Transfer Site MTD Sold Qty': transfer.get('Transfer Site MTD Sold Qty', 0),
                    'Receive Site Last Month Sold Qty': transfer.get('Receive Site Last Month Sold Qty', 0),
                    'Receive Site MTD Sold Qty': transfer.get('Receive Site MTD Sold Qty', 0),
                    'Receive Original Stock': receive_original_stock
                }
                transfer_data.append(row)
            
            transfer_df = pd.DataFrame(transfer_data)
            transfer_df.to_excel(writer, sheet_name='調貨建議', index=False)

        # 工作表2: 統計摘要
        # 基本KPI
        basic_stats = pd.DataFrame([{
            '指標': '涉及行數',
            '數值': stats['basic']['total_recommendations']
        }, {
            '指標': '總轉移量',
            '數值': stats['basic']['total_transfer_qty']
        }, {
            '指標': '涉及SKU數量',
            '數值': stats['basic']['unique_articles']
        }, {
            '指標': '涉及OM',
            '數值': stats['basic']['unique_oms']
        }])

        start_row = 0
        basic_stats.to_excel(writer, sheet_name='統計摘要', startrow=start_row, index=False)

        # 按商品統計
        start_row += len(basic_stats) + 3
        if stats['by_article']:
            article_df = pd.DataFrame(stats['by_article'])
            article_df.to_excel(writer, sheet_name='統計摘要', startrow=start_row, index=False)

        # 按OM統計
        start_row += len(stats['by_article']) + 3
        if stats['by_om']:
            om_df = pd.DataFrame(stats['by_om'])
            om_df.to_excel(writer, sheet_name='統計摘要', startrow=start_row, index=False)

        # 轉移類型
        start_row += len(stats['by_om']) + 3
        if stats['transfer_types']:
            type_data = []
            for ttype, data in stats['transfer_types'].items():
                type_data.append({
                    '轉移類型': ttype,
                    '總量': data['qty'],
                    '行數': data['lines']
                })
            type_df = pd.DataFrame(type_data)
            type_df.to_excel(writer, sheet_name='統計摘要', startrow=start_row, index=False)

        # 接收統計
        start_row += len(type_data) + 3
        if stats['receive_stats']:
            receive_df = pd.DataFrame(stats['receive_stats'])
            receive_df.to_excel(writer, sheet_name='統計摘要', startrow=start_row, index=False)

    output.seek(0)
    return output

# Main UI
st.header("1. 資料上傳")
uploaded_file = st.file_uploader("上傳Excel檔案", type=['xlsx', 'xls'])

if uploaded_file is not None:
    with st.spinner("正在載入並驗證資料..."):
        data = load_data(uploaded_file)
        if data is not None:
            st.session_state.data = data
            st.success(f"資料載入成功！共處理 {len(data)} 行資料。")

            # 資料預覽
            st.header("2. 資料預覽")
            st.subheader("樣本資料")
            st.dataframe(data.head(10))

            st.subheader("基本統計")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("總行數", len(data))
            with col2:
                st.metric("唯一商品數", data['Article'].nunique())
            with col3:
                st.metric("唯一店舖數", data['Site'].nunique())
            with col4:
                st.metric("唯一OM組別數", data['OM'].nunique())

            # 預處理資料
            with st.spinner("正在預處理資料..."):
                processed_data = preprocess_data(data)
                st.session_state.processed_data = processed_data

            # 模式選擇
            st.header("3. 轉移模式選擇")
            
            # 顯示模式特性說明
            with st.expander("📊 查看各模式特性說明", expanded=True):
                st.markdown("""
                **模式選擇指南：**
                
                #### **A: 保守轉移** - 適用於穩定期
                - **ND類型**：完整轉移所有庫存
                - **RF類型**：僅轉移超出安全庫存的部分，最多轉移50%
                - **OM限制**：僅限相同OM組別內調撥
                - **適用場景**：庫存充足，需要謹慎調撥
                
                #### **B: 增強轉移** - 適用於成長期
                - **ND類型**：完整轉移所有庫存
                - **RF類型**：僅轉移超出MOQ的部分，最多轉移90%
                - **OM限制**：僅限相同OM組別內調撥
                - **適用場景**：需要積極調撥，但保留部分庫存
                
                #### **C: 超級增強轉移** - 適用於緊急調撥
                - **ND類型**：完整轉移所有庫存
                - **RF類型**：可轉移全部庫存（無保留限制）
                - **OM限制**：**允許跨組調撥**（HD組別除外）
                - **適用場景**：緊急需求，需要最大化調撥
                
                **OM調撥限制（公司現有組別：HA, HB, HC, HD, HZ）：**
                - ✅ **可調撥**：HA ↔ HB, HC, HZ | HB ↔ HA, HC, HZ | HC ↔ HA, HB, HZ | HZ ↔ HA, HB, HC
                - ❌ **不可調撥**：HD → HA, HB, HC
                - ✅ **可調撥**：HD → HZ
                """)
            
            mode = st.radio(
                "選擇轉移模式:",
                ['A: 保守轉移', 'B: 增強轉移', 'C: 超級增強轉移'],
                index=0
            )
            st.session_state.mode = mode[0]

            # 生成建議
            if st.button("生成轉移建議", type="primary"):
                with st.spinner("正在生成建議..."):
                    if mode.startswith('A'):
                        transfers = generate_transfer_recommendations_conservative(processed_data)
                    elif mode.startswith('B'):
                        transfers = generate_transfer_recommendations_enhanced(processed_data)
                    else:  # Mode C
                        transfers = generate_transfer_recommendations_super(processed_data)

                    st.session_state.transfer_results = transfers

                    if transfers:
                        st.success(f"成功生成 {len(transfers)} 條轉移建議！")

                        # 統計分析
                        st.header("4. 分析結果")
                        stats = calculate_statistics(transfers, processed_data)

                        # KPI 卡片
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("涉及行數", stats['basic']['total_recommendations'])
                        with col2:
                            st.metric("總轉移量", stats['basic']['total_transfer_qty'])
                        with col3:
                            st.metric("涉及SKU數量", stats['basic']['unique_articles'])
                        with col4:
                            st.metric("涉及OM", stats['basic']['unique_oms'])

                        # 轉移結果表格
                        st.subheader("轉移建議明細")
                        transfer_df = pd.DataFrame(transfers)
                        st.dataframe(transfer_df)

                        # 統計表格
                        st.subheader("按商品統計")
                        if stats['by_article']:
                            st.dataframe(pd.DataFrame(stats['by_article']))

                        st.subheader("按OM統計")
                        if stats['by_om']:
                            st.dataframe(pd.DataFrame(stats['by_om']))

                        st.subheader("轉移類型分佈")
                        if stats['transfer_types']:
                            type_data = []
                            for ttype, data in stats['transfer_types'].items():
                                type_data.append({
                                    '類型': ttype,
                                    '總量': data['qty'],
                                    '行數': data['lines']
                                })
                            st.dataframe(pd.DataFrame(type_data))

                        # 視覺化
                        st.subheader("轉移分析圖表")
                        fig = create_visualization(transfers, st.session_state.mode, processed_data)
                        if fig:
                            st.pyplot(fig)

                        # 匯出
                        st.header("5. 匯出結果")
                        excel_data = export_to_excel(transfers, stats, processed_data)
                        st.download_button(
                            label="📥 下載Excel檔案",
                            data=excel_data.getvalue(),
                            file_name=f"店舖轉移建議_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="excel_download"
                        )
                    else:
                        st.warning("未生成轉移建議。請檢查您的資料並嘗試不同模式。")

else:
    st.info("請上傳Excel檔案開始。")

# Footer
st.markdown("---")
st.markdown("*開發者: Ricky - 版本 2.0*")