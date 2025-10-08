import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import warnings
warnings.filterwarnings('ignore')

# Set page configuration
st.set_page_config(
    page_title="調貨建議生成系統",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .sidebar-header {
        font-size: 1.2rem;
        font-weight: bold;
        color: #2c3e50;
        margin-bottom: 1rem;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 0.25rem solid #1f77b4;
    }
</style>
""", unsafe_allow_html=True)

class InventoryTransferSystem:
    def __init__(self):
        self.df = None
        self.transfer_suggestions = None
        self.analysis_results = None

    def load_and_preprocess_data(self, file):
        """Load and preprocess Excel data according to specifications"""
        try:
            # Read Excel file
            df = pd.read_excel(file)

            # Validate required columns
            required_columns = [
                'Article', 'Article Description', 'RP Type', 'Site', 'OM', 'MOQ',
                'SaSa Net Stock', 'Target', 'Pending Received', 'Safety Stock',
                'Last Month Sold Qty', 'MTD Sold Qty'
            ]

            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"缺少必需欄位: {', '.join(missing_columns)}")

            # Data preprocessing
            df = df.copy()

            # 1. Convert Article to string
            df['Article'] = df['Article'].astype(str)

            # 2. Convert numeric columns, fill invalid values with 0
            numeric_columns = [
                'MOQ', 'SaSa Net Stock', 'Target', 'Pending Received',
                'Safety Stock', 'Last Month Sold Qty', 'MTD Sold Qty'
            ]

            for col in numeric_columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

            # 3. Fix negative values
            for col in numeric_columns:
                df[col] = df[col].clip(lower=0)

            # 4. Handle sales outliers
            for col in ['Last Month Sold Qty', 'MTD Sold Qty']:
                outlier_mask = df[col] > 100000
                if outlier_mask.any():
                    df.loc[outlier_mask, col] = 100000
                    df.loc[outlier_mask, 'Notes'] = df.loc[outlier_mask, 'Notes'].fillna('') + '銷量異常值已調整; '

            # 5. Fill string columns
            string_columns = ['Article Description', 'RP Type', 'Site', 'OM']
            for col in string_columns:
                df[col] = df[col].fillna('').astype(str)

            # 6. Add Notes column for data cleaning logs
            df['Notes'] = ''

            # Validate RP Type values
            invalid_rp_types = df[~df['RP Type'].isin(['ND', 'RF'])]
            if not invalid_rp_types.empty:
                invalid_sites = invalid_rp_types['Site'].unique()
                df.loc[df['Site'].isin(invalid_sites), 'Notes'] += 'RP Type值無效; '

            self.df = df
            return True, "資料載入成功"

        except Exception as e:
            return False, f"資料載入失敗: {str(e)}"

    def calculate_transfer_suggestions(self, strategy='A'):
        """Calculate transfer suggestions based on selected strategy"""
        if self.df is None:
            return False, "請先載入資料"

        try:
            df = self.df.copy()

            # Calculate effective sales
            df['Effective_Sales'] = np.where(
                df['Last Month Sold Qty'] > 0,
                df['Last Month Sold Qty'],
                df['MTD Sold Qty']
            )

            # Get max sales for each product within same OM
            df['Max_Sales_In_OM'] = df.groupby(['Article', 'OM'])['Effective_Sales'].transform('max')

            # Strategy A: Conservative Transfer
            if strategy == 'A':
                return self._calculate_strategy_a(df)
            # Strategy B: Enhanced Transfer
            elif strategy == 'B':
                return self._calculate_strategy_b(df)
            # Strategy C: Super Enhanced Transfer
            elif strategy == 'C':
                return self._calculate_strategy_c(df)
            else:
                return False, "無效的策略選擇"

        except Exception as e:
            return False, f"計算失敗: {str(e)}"

    def _calculate_strategy_a(self, df):
        """Strategy A: Conservative Transfer"""
        transfer_candidates = []
        receive_candidates = []

        # 1. ND Type - Complete transfer out
        nd_candidates = df[
            (df['RP Type'] == 'ND') &
            (df['SaSa Net Stock'] > 0)
        ].copy()

        for _, row in nd_candidates.iterrows():
            transfer_candidates.append({
                'Article': row['Article'],
                'Product_Desc': row['Article Description'],
                'OM': row['OM'],
                'Transfer_Site': row['Site'],
                'Transfer_Qty': int(row['SaSa Net Stock']),
                'Transfer_Site_Original_Stock': int(row['SaSa Net Stock']),
                'Transfer_Site_After_Transfer_Stock': 0,
                'Transfer_Site_Safety_Stock': int(row['Safety Stock']),
                'Transfer_Site_MOQ': int(row['MOQ']),
                'Transfer_Type': 'ND轉出',
                'Priority': 1,
                'Notes': row.get('Notes', '')
            })

        # 2. RF Type - Excess transfer out
        rf_candidates = df[
            (df['RP Type'] == 'RF') &
            ((df['SaSa Net Stock'] + df['Pending Received']) > df['Safety Stock']) &
            (df['Effective_Sales'] < df['Max_Sales_In_OM'])
        ].copy()

        # Sort by effective sales (lowest first for transfer out)
        rf_candidates = rf_candidates.sort_values('Effective_Sales')

        for _, row in rf_candidates.iterrows():
            current_stock = row['SaSa Net Stock'] + row['Pending Received']
            safety_stock = row['Safety Stock']
            base_transferable = current_stock - safety_stock
            max_transferable = int(current_stock * 0.5)
            actual_transfer = min(base_transferable, max_transferable, row['SaSa Net Stock'])

            if actual_transfer > 0:
                transfer_candidates.append({
                    'Article': row['Article'],
                    'Product_Desc': row['Article Description'],
                    'OM': row['OM'],
                    'Transfer_Site': row['Site'],
                    'Transfer_Qty': actual_transfer,
                    'Transfer_Site_Original_Stock': int(row['SaSa Net Stock']),
                    'Transfer_Site_After_Transfer_Stock': int(row['SaSa Net Stock'] - actual_transfer),
                    'Transfer_Site_Safety_Stock': int(row['Safety Stock']),
                    'Transfer_Site_MOQ': int(row['MOQ']),
                    'Transfer_Type': 'RF過剩轉出',
                    'Priority': 2,
                    'Notes': row.get('Notes', '')
                })

        # 3. Receive candidates - Sites with Target quantity
        receive_candidates_df = df[df['Target'] > 0].copy()

        for _, row in receive_candidates_df.iterrows():
            receive_candidates.append({
                'Article': row['Article'],
                'Product_Desc': row['Article Description'],
                'OM': row['OM'],
                'Receive_Site': row['Site'],
                'Target_Qty': int(row['Target']),
                'Receive_Type': '指定店鋪補貨',
                'Priority': 1,
                'Notes': row.get('Notes', '')
            })

        # 4. Match transfers with receives
        suggestions = self._match_transfers_to_receives(transfer_candidates, receive_candidates)

        self.transfer_suggestions = pd.DataFrame(suggestions)
        return True, f"策略A計算完成，共產生 {len(suggestions)} 筆建議"

    def _calculate_strategy_b(self, df):
        """Strategy B: Enhanced Transfer"""
        transfer_candidates = []
        receive_candidates = []

        # 1. ND Type - Complete transfer out (same as Strategy A)
        nd_candidates = df[
            (df['RP Type'] == 'ND') &
            (df['SaSa Net Stock'] > 0)
        ].copy()

        for _, row in nd_candidates.iterrows():
            transfer_candidates.append({
                'Article': row['Article'],
                'Product_Desc': row['Article Description'],
                'OM': row['OM'],
                'Transfer_Site': row['Site'],
                'Transfer_Qty': int(row['SaSa Net Stock']),
                'Transfer_Site_Original_Stock': int(row['SaSa Net Stock']),
                'Transfer_Site_After_Transfer_Stock': 0,
                'Transfer_Site_Safety_Stock': int(row['Safety Stock']),
                'Transfer_Site_MOQ': int(row['MOQ']),
                'Transfer_Type': 'ND轉出',
                'Priority': 1,
                'Notes': row.get('Notes', '')
            })

        # 2. RF Type - Enhanced transfer out
        rf_candidates = df[
            (df['RP Type'] == 'RF') &
            ((df['SaSa Net Stock'] + df['Pending Received']) > (df['MOQ'] + 1)) &
            (df['Effective_Sales'] < df['Max_Sales_In_OM'])
        ].copy()

        # Sort by effective sales (lowest first for transfer out)
        rf_candidates = rf_candidates.sort_values('Effective_Sales')

        for _, row in rf_candidates.iterrows():
            current_stock = row['SaSa Net Stock'] + row['Pending Received']
            moq = row['MOQ']
            base_transferable = current_stock - (moq + 1)
            max_transferable = int(current_stock * 0.8)
            actual_transfer = min(base_transferable, max_transferable, row['SaSa Net Stock'])

            if actual_transfer > 0:
                transfer_candidates.append({
                    'Article': row['Article'],
                    'Product_Desc': row['Article Description'],
                    'OM': row['OM'],
                    'Transfer_Site': row['Site'],
                    'Transfer_Qty': actual_transfer,
                    'Transfer_Site_Original_Stock': int(row['SaSa Net Stock']),
                    'Transfer_Site_After_Transfer_Stock': int(row['SaSa Net Stock'] - actual_transfer),
                    'Transfer_Site_Safety_Stock': int(row['Safety Stock']),
                    'Transfer_Site_MOQ': int(row['MOQ']),
                    'Transfer_Type': 'RF加強轉出',
                    'Priority': 2,
                    'Notes': row.get('Notes', '')
                })

        # 3. Receive candidates (same as Strategy A)
        receive_candidates_df = df[df['Target'] > 0].copy()

        for _, row in receive_candidates_df.iterrows():
            receive_candidates.append({
                'Article': row['Article'],
                'Product_Desc': row['Article Description'],
                'OM': row['OM'],
                'Receive_Site': row['Site'],
                'Target_Qty': int(row['Target']),
                'Receive_Type': '指定店鋪補貨',
                'Priority': 1,
                'Notes': row.get('Notes', '')
            })

        # 4. Match transfers with receives
        suggestions = self._match_transfers_to_receives(transfer_candidates, receive_candidates)

        self.transfer_suggestions = pd.DataFrame(suggestions)
        return True, f"策略B計算完成，共產生 {len(suggestions)} 筆建議"

    def _calculate_strategy_c(self, df):
        """Strategy C: Super Enhanced Transfer"""
        transfer_candidates = []
        receive_candidates = []

        # 1. ND Type - Complete transfer out (same as Strategy A)
        nd_candidates = df[
            (df['RP Type'] == 'ND') &
            (df['SaSa Net Stock'] > 0)
        ].copy()

        for _, row in nd_candidates.iterrows():
            transfer_candidates.append({
                'Article': row['Article'],
                'Product_Desc': row['Article Description'],
                'OM': row['OM'],
                'Transfer_Site': row['Site'],
                'Transfer_Qty': int(row['SaSa Net Stock']),
                'Transfer_Site_Original_Stock': int(row['SaSa Net Stock']),
                'Transfer_Site_After_Transfer_Stock': 0,
                'Transfer_Site_Safety_Stock': int(row['Safety Stock']),
                'Transfer_Site_MOQ': int(row['MOQ']),
                'Transfer_Type': 'ND轉出',
                'Priority': 1,
                'Notes': row.get('Notes', '')
            })

        # 2. RF Type - Super enhanced transfer out
        rf_candidates = df[
            (df['RP Type'] == 'RF') &
            ((df['SaSa Net Stock'] + df['Pending Received']) > 0) &
            (df['Effective_Sales'] < df['Max_Sales_In_OM'])
        ].copy()

        # Sort by effective sales (lowest first for transfer out)
        rf_candidates = rf_candidates.sort_values('Effective_Sales')

        for _, row in rf_candidates.iterrows():
            current_stock = row['SaSa Net Stock'] + row['Pending Received']
            base_transferable = max(0, current_stock - 2)  # Leave 2 pieces at most
            max_transferable = int(current_stock * 0.9)
            actual_transfer = min(base_transferable, max_transferable, row['SaSa Net Stock'])

            if actual_transfer > 0:
                transfer_candidates.append({
                    'Article': row['Article'],
                    'Product_Desc': row['Article Description'],
                    'OM': row['OM'],
                    'Transfer_Site': row['Site'],
                    'Transfer_Qty': actual_transfer,
                    'Transfer_Site_Original_Stock': int(row['SaSa Net Stock']),
                    'Transfer_Site_After_Transfer_Stock': int(row['SaSa Net Stock'] - actual_transfer),
                    'Transfer_Site_Safety_Stock': int(row['Safety Stock']),
                    'Transfer_Site_MOQ': int(row['MOQ']),
                    'Transfer_Type': 'RF特強轉出',
                    'Priority': 2,
                    'Notes': row.get('Notes', '')
                })

        # 3. Receive candidates (same as Strategy A)
        receive_candidates_df = df[df['Target'] > 0].copy()

        for _, row in receive_candidates_df.iterrows():
            receive_candidates.append({
                'Article': row['Article'],
                'Product_Desc': row['Article Description'],
                'OM': row['OM'],
                'Receive_Site': row['Site'],
                'Target_Qty': int(row['Target']),
                'Receive_Type': '指定店鋪補貨',
                'Priority': 1,
                'Notes': row.get('Notes', '')
            })

        # 4. Match transfers with receives
        suggestions = self._match_transfers_to_receives(transfer_candidates, receive_candidates)

        self.transfer_suggestions = pd.DataFrame(suggestions)
        return True, f"策略C計算完成，共產生 {len(suggestions)} 筆建議"

    def _match_transfers_to_receives(self, transfer_candidates, receive_candidates):
        """Match transfer candidates with receive candidates"""
        suggestions = []

        # Sort by priority
        transfer_candidates.sort(key=lambda x: x['Priority'])
        receive_candidates.sort(key=lambda x: x['Priority'])

        # Group by Article and OM for matching
        transfer_by_article_om = {}
        for candidate in transfer_candidates:
            key = (candidate['Article'], candidate['OM'])
            if key not in transfer_by_article_om:
                transfer_by_article_om[key] = []
            transfer_by_article_om[key].append(candidate)

        receive_by_article_om = {}
        for candidate in receive_candidates:
            key = (candidate['Article'], candidate['OM'])
            if key not in receive_by_article_om:
                receive_by_article_om[key] = []
            receive_by_article_om[key].append(candidate)

        # Match transfers to receives
        for (article, om), transfers in transfer_by_article_om.items():
            receives = receive_by_article_om.get((article, om), [])

            for transfer in transfers:
                remaining_qty = transfer['Transfer_Qty']

                for receive in receives:
                    if remaining_qty <= 0:
                        break

                    # Avoid same site transfer
                    if transfer['Transfer_Site'] == receive['Receive_Site']:
                        continue

                    transfer_qty = min(remaining_qty, receive['Target_Qty'])

                    if transfer_qty > 0:
                        suggestions.append({
                            'Article': transfer['Article'],
                            'Product_Desc': transfer['Product_Desc'],
                            'OM': transfer['OM'],
                            'Transfer_Site': transfer['Transfer_Site'],
                            'Transfer_Qty': transfer_qty,
                            'Transfer_Site_Original_Stock': transfer['Transfer_Site_Original_Stock'],
                            'Transfer_Site_After_Transfer_Stock': transfer['Transfer_Site_After_Transfer_Stock'],
                            'Transfer_Site_Safety_Stock': transfer['Transfer_Site_Safety_Stock'],
                            'Transfer_Site_MOQ': transfer['Transfer_Site_MOQ'],
                            'Receive_Site': receive['Receive_Site'],
                            'Receive_Site_Target_Qty': receive['Target_Qty'],
                            'Actual_Receive_Qty': transfer_qty,
                            'Transfer_Type': transfer['Transfer_Type'],
                            'Notes': transfer['Notes'] + receive['Notes']
                        })

                        remaining_qty -= transfer_qty
                        receive['Target_Qty'] -= transfer_qty

        return suggestions

    def generate_visualization(self, strategy):
        """Generate matplotlib horizontal bar chart"""
        if self.transfer_suggestions is None:
            return None

        # Group by OM and transfer type
        om_summary = self.transfer_suggestions.groupby(['OM', 'Transfer_Type'])['Transfer_Qty'].sum().unstack(fill_value=0)

        # Define transfer types based on strategy
        if strategy == 'A':
            transfer_types = ['ND轉出', 'RF過剩轉出']
        elif strategy == 'B':
            transfer_types = ['ND轉出', 'RF加強轉出']
        else:  # Strategy C
            transfer_types = ['ND轉出', 'RF特強轉出']

        # Ensure all transfer types are present
        for t_type in transfer_types:
            if t_type not in om_summary.columns:
                om_summary[t_type] = 0

        # Calculate receive quantities
        receive_summary = self.transfer_suggestions.groupby('OM')['Actual_Receive_Qty'].sum()

        # Combine all data
        plot_data = []
        om_list = sorted(om_summary.index)

        for om in om_list:
            plot_data.append({
                'OM': om,
                'ND轉出': om_summary.loc[om, 'ND轉出'] if 'ND轉出' in om_summary.columns else 0,
                'RF過剩轉出': om_summary.loc[om, 'RF過剩轉出'] if 'RF過剩轉出' in om_summary.columns else 0,
                'RF加強轉出': om_summary.loc[om, 'RF加強轉出'] if 'RF加強轉出' in om_summary.columns else 0,
                'RF特強轉出': om_summary.loc[om, 'RF特強轉出'] if 'RF特強轉出' in om_summary.columns else 0,
                '需求接收數量': self.transfer_suggestions[self.transfer_suggestions['OM'] == om]['Receive_Site_Target_Qty'].sum(),
                '實際接收數量': receive_summary.get(om, 0)
            })

        # Create the plot
        fig, ax = plt.subplots(figsize=(12, 8))

        bar_width = 0.15
        index = np.arange(len(om_list))

        # Plot bars for each category
        bars = []
        if strategy == 'A':
            bars.append(ax.barh(index - bar_width*1.5, [d['ND轉出'] for d in plot_data], bar_width, label='ND轉出', alpha=0.8))
            bars.append(ax.barh(index - bar_width*0.5, [d['RF過剩轉出'] for d in plot_data], bar_width, label='RF過剩轉出', alpha=0.8))
            bars.append(ax.barh(index + bar_width*0.5, [d['需求接收數量'] for d in plot_data], bar_width, label='需求接收數量', alpha=0.8))
            bars.append(ax.barh(index + bar_width*1.5, [d['實際接收數量'] for d in plot_data], bar_width, label='實際接收數量', alpha=0.8))
        elif strategy == 'B':
            bars.append(ax.barh(index - bar_width*2, [d['ND轉出'] for d in plot_data], bar_width, label='ND轉出', alpha=0.8))
            bars.append(ax.barh(index - bar_width, [d['RF加強轉出'] for d in plot_data], bar_width, label='RF加強轉出', alpha=0.8))
            bars.append(ax.barh(index, [d['需求接收數量'] for d in plot_data], bar_width, label='需求接收數量', alpha=0.8))
            bars.append(ax.barh(index + bar_width, [d['實際接收數量'] for d in plot_data], bar_width, label='實際接收數量', alpha=0.8))
        else:  # Strategy C
            bars.append(ax.barh(index - bar_width*2, [d['ND轉出'] for d in plot_data], bar_width, label='ND轉出', alpha=0.8))
            bars.append(ax.barh(index - bar_width, [d['RF特強轉出'] for d in plot_data], bar_width, label='RF特強轉出', alpha=0.8))
            bars.append(ax.barh(index, [d['需求接收數量'] for d in plot_data], bar_width, label='需求接收數量', alpha=0.8))
            bars.append(ax.barh(index + bar_width, [d['實際接收數量'] for d in plot_data], bar_width, label='實際接收數量', alpha=0.8))

        # Customize the plot
        ax.set_xlabel('數量', fontsize=12)
        ax.set_ylabel('OM單位', fontsize=12)
        ax.set_title('調貨接收分析', fontsize=14, fontweight='bold')
        ax.set_yticks(index)
        ax.set_yticklabels(om_list)
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
        ax.grid(axis='x', alpha=0.3)

        # Add value labels on bars
        for bar in bars:
            for rect in bar:
                width = rect.get_width()
                if width > 0:
                    ax.text(width + max([d['ND轉出'] + d['RF過剩轉出'] + d['RF加強轉出'] + d['RF特強轉出'] + d['需求接收數量'] + d['實際接收數量'] for d in plot_data]) * 0.01,
                           rect.get_y() + rect.get_height()/2,
                           f'{int(width)}',
                           ha='left', va='center', fontsize=9)

        plt.tight_layout()
        return fig

    def generate_statistics(self):
        """Generate comprehensive statistics"""
        if self.transfer_suggestions is None:
            return None

        df = self.transfer_suggestions

        # Basic KPIs
        total_suggestions = len(df)
        total_transfer_qty = df['Transfer_Qty'].sum()
        unique_products = df['Article'].nunique()
        unique_oms = df['OM'].nunique()

        # Statistics by product
        product_stats = df.groupby('Article').agg({
            'Target_Qty': 'sum',
            'Transfer_Qty': 'sum',
            'Article': 'size',
            'Receive_Site_Target_Qty': 'sum'
        }).rename(columns={
            'Article': 'Transfer_Count',
            'Target_Qty': 'Total_Target_Qty',
            'Receive_Site_Target_Qty': 'Total_Receive_Target_Qty'
        })

        product_stats['Fulfillment_rate'] = (product_stats['Total_Transfer_Qty'] / product_stats['Total_Receive_Target_Qty'] * 100).round(2)

        # Statistics by OM
        om_stats = df.groupby('OM').agg({
            'Transfer_Qty': 'sum',
            'Target_Qty': 'sum',
            'Article': 'size',
            'Receive_Site_Target_Qty': 'sum'
        }).rename(columns={
            'Article': 'Transfer_Count',
            'Target_Qty': 'Total_Target_Qty',
            'Receive_Site_Target_Qty': 'Total_Receive_Target_Qty'
        })

        om_stats['Unique_Products'] = df.groupby('OM')['Article'].nunique()

        # Transfer type distribution
        transfer_type_stats = df.groupby('Transfer_Type').agg({
            'Transfer_Qty': 'sum',
            'Transfer_Type': 'size'
        }).rename(columns={
            'Transfer_Type': 'Count'
        })

        # Receive type distribution
        receive_summary = {
            'Total_Target_Qty': df['Receive_Site_Target_Qty'].sum(),
            'Total_Actual_Receive_Qty': df['Actual_Receive_Qty'].sum(),
            'Fulfillment_rate': (df['Actual_Receive_Qty'].sum() / df['Receive_Site_Target_Qty'].sum() * 100) if df['Receive_Site_Target_Qty'].sum() > 0 else 0
        }

        return {
            'basic_kpis': {
                'total_suggestions': total_suggestions,
                'total_transfer_qty': total_transfer_qty,
                'unique_products': unique_products,
                'unique_oms': unique_oms
            },
            'product_stats': product_stats,
            'om_stats': om_stats,
            'transfer_type_stats': transfer_type_stats,
            'receive_summary': receive_summary
        }

    def export_to_excel(self):
        """Export results to Excel with multiple worksheets"""
        if self.transfer_suggestions is None:
            return None

        # Create workbook
        wb = Workbook()

        # 1. Transfer Suggestions worksheet
        ws1 = wb.active
        ws1.title = "調貨建議"

        # Headers
        headers = [
            'Article', 'Product Desc', 'OM', 'Transfer Site', 'Transfer Qty',
            'Transfer Site Original Stock', 'Transfer Site After Transfer Stock',
            'Transfer Site Safety Stock', 'Transfer Site MOQ', 'Receive Site',
            'Receive Site Target Qty', 'Notes'
        ]

        for col, header in enumerate(headers, 1):
            cell = ws1.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True)

        # Data
        for row, (_, record) in enumerate(self.transfer_suggestions.iterrows(), 2):
            ws1.cell(row=row, column=1, value=record['Article'])
            ws1.cell(row=row, column=2, value=record['Product_Desc'])
            ws1.cell(row=row, column=3, value=record['OM'])
            ws1.cell(row=row, column=4, value=record['Transfer_Site'])
            ws1.cell(row=row, column=5, value=record['Transfer_Qty'])
            ws1.cell(row=row, column=6, value=record['Transfer_Site_Original_Stock'])
            ws1.cell(row=row, column=7, value=record['Transfer_Site_After_Transfer_Stock'])
            ws1.cell(row=row, column=8, value=record['Transfer_Site_Safety_Stock'])
            ws1.cell(row=row, column=9, value=record['Transfer_Site_MOQ'])
            ws1.cell(row=row, column=10, value=record['Receive_Site'])
            ws1.cell(row=row, column=11, value=record['Receive_Site_Target_Qty'])
            ws1.cell(row=row, column=12, value=record['Notes'])

        # 2. Statistics Summary worksheet
        ws2 = wb.create_worksheet("統計摘要")

        stats = self.generate_statistics()
        if stats:
            current_row = 1

            # Basic KPIs
            ws2.cell(row=current_row, column=1, value="基本KPI指標").font = Font(bold=True)
            current_row += 2

            kpi_data = [
                ['總調貨建議數量', stats['basic_kpis']['total_suggestions']],
                ['總調貨件數', stats['basic_kpis']['total_transfer_qty']],
                ['涉及產品數量', stats['basic_kpis']['unique_products']],
                ['涉及OM數量', stats['basic_kpis']['unique_oms']]
            ]

            for col, (label, value) in enumerate(kpi_data, 1):
                ws2.cell(row=current_row, column=col, value=label).font = Font(bold=True)
                ws2.cell(row=current_row + 1, column=col, value=value)

            current_row += 4

            # Product Statistics
            ws2.cell(row=current_row, column=1, value="按產品統計").font = Font(bold=True)
            current_row += 2

            product_headers = ['產品編號', '總需求件數', '總調貨件數', '調貨行數', 'Fullfillment Rate (%)']
            for col, header in enumerate(product_headers, 1):
                ws2.cell(row=current_row, column=col, value=header).font = Font(bold=True)

            current_row += 1

            for article, row in stats['product_stats'].iterrows():
                ws2.cell(row=current_row, column=1, value=article)
                ws2.cell(row=current_row, column=2, value=row['Total_Receive_Target_Qty'])
                ws2.cell(row=current_row, column=3, value=row['Total_Transfer_Qty'])
                ws2.cell(row=current_row, column=4, value=row['Transfer_Count'])
                ws2.cell(row=current_row, column=5, value=row['fulfillment_rate'])
                current_row += 1

            current_row += 3

            # OM Statistics
            ws2.cell(row=current_row, column=1, value="按OM統計").font = Font(bold=True)
            current_row += 2

            om_headers = ['OM單位', '總調貨件數', '總需求件數', '調貨行數', '涉及產品數量']
            for col, header in enumerate(om_headers, 1):
                ws2.cell(row=current_row, column=col, value=header).font = Font(bold=True)

            current_row += 1

            for om, row in stats['om_stats'].iterrows():
                ws2.cell(row=current_row, column=1, value=om)
                ws2.cell(row=current_row, column=2, value=row['Transfer_Qty'])
                ws2.cell(row=current_row, column=3, value=row['Total_Receive_Target_Qty'])
                ws2.cell(row=current_row, column=4, value=row['Transfer_Count'])
                ws2.cell(row=current_row, column=5, value=row['Unique_Products'])
                current_row += 1

            current_row += 3

            # Transfer Type Distribution
            ws2.cell(row=current_row, column=1, value="轉出類型分佈").font = Font(bold=True)
            current_row += 2

            transfer_headers = ['轉出類型', '總件數', '涉及行數']
            for col, header in enumerate(transfer_headers, 1):
                ws2.cell(row=current_row, column=col, value=header).font = Font(bold=True)

            current_row += 1

            for t_type, row in stats['transfer_type_stats'].iterrows():
                ws2.cell(row=current_row, column=1, value=t_type)
                ws2.cell(row=current_row, column=2, value=row['Transfer_Qty'])
                ws2.cell(row=current_row, column=3, value=row['Count'])
                current_row += 1

            current_row += 3

            # Receive Summary
            ws2.cell(row=current_row, column=1, value="接收類型結果").font = Font(bold=True)
            current_row += 2

            receive_data = [
                ['總需求數量', stats['receive_summary']['Total_Target_Qty']],
                ['總實際接收數量', stats['receive_summary']['Total_Actual_Receive_Qty']],
                ['達成率 (%)', stats['receive_summary']['fulfillment_rate']]
            ]

            for col, (label, value) in enumerate(receive_data, 1):
                ws2.cell(row=current_row, column=col, value=label).font = Font(bold=True)
                ws2.cell(row=current_row + 1, column=col, value=value)

        # Generate filename with current date
        current_date = datetime.now().strftime('%Y%m%d')
        filename = f'強制轉貨建議_{current_date}.xlsx'

        # Save to BytesIO
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)

        return excel_buffer, filename

# Initialize the system
system = InventoryTransferSystem()

# Main UI
def main():
    # Header
    st.markdown('<div class="main-header">📦 調貨建議生成系統</div>', unsafe_allow_html=True)
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
    st.subheader("1. 資料上傳區塊")
    uploaded_file = st.file_uploader(
        "請選擇Excel檔案",
        type=['xlsx', 'xls'],
        help="支援 .xlsx 和 .xls 格式，檔案必須包含指定的欄位"
    )

    if uploaded_file is not None:
        with st.spinner('載入資料中...'):
            success, message = system.load_and_preprocess_data(uploaded_file)

        if success:
            st.success(message)

            # Data preview section
            st.subheader("2. 資料預覽區塊")
            col1, col2 = st.columns([2, 1])

            with col1:
                st.write("資料樣本預覽:")
                st.dataframe(system.df.head())

            with col2:
                st.write("基本統計資訊:")
                st.info(f"""
                總記錄數: {len(system.df)}
                產品數量: {system.df['Article'].nunique()}
                店鋪數量: {system.df['Site'].nunique()}
                OM數量: {system.df['OM'].nunique()}
                """)

            # Strategy selection and analysis
            st.subheader("3. 分析按鈕區塊")
            st.write("請選擇轉貨策略:")

            col1, col2, col3 = st.columns(3)

            with col1:
                if st.button("🛡️ A: 保守轉貨", type="primary", use_container_width=True):
                    with st.spinner('計算轉貨建議中...'):
                        success, message = system.calculate_transfer_suggestions('A')
                    if success:
                        st.success(message)
                        st.session_state.analysis_complete = True
                        st.session_state.strategy = 'A'
                    else:
                        st.error(message)

            with col2:
                if st.button("⚡ B: 加強轉貨", type="primary", use_container_width=True):
                    with st.spinner('計算轉貨建議中...'):
                        success, message = system.calculate_transfer_suggestions('B')
                    if success:
                        st.success(message)
                        st.session_state.analysis_complete = True
                        st.session_state.strategy = 'B'
                    else:
                        st.error(message)

            with col3:
                if st.button("🚀 C: 特強轉貨", type="primary", use_container_width=True):
                    with st.spinner('計算轉貨建議中...'):
                        success, message = system.calculate_transfer_suggestions('C')
                    if success:
                        st.success(message)
                        st.session_state.analysis_complete = True
                        st.session_state.strategy = 'C'
                    else:
                        st.error(message)

        else:
            st.error(message)

    # Results section
    if st.session_state.get('analysis_complete', False):
        st.subheader("4. 結果展示區塊")

        # KPI metrics
        stats = system.generate_statistics()
        if stats:
            st.write("📊 KPI指標卡:")
            col1, col2, col3, col4 = st.columns(4)

            with col1:
                st.metric(
                    label="總調貨建議數量",
                    value=stats['basic_kpis']['total_suggestions']
                )

            with col2:
                st.metric(
                    label="總調貨件數",
                    value=stats['basic_kpis']['total_transfer_qty']
                )

            with col3:
                st.metric(
                    label="涉及產品數量",
                    value=stats['basic_kpis']['unique_products']
                )

            with col4:
                st.metric(
                    label="涉及OM數量",
                    value=stats['basic_kpis']['unique_oms']
                )

        # Transfer suggestions table
        st.write("📋 調貨建議表格:")
        st.dataframe(system.transfer_suggestions)

        # Visualization
        st.write("📈 統計圖表:")
        fig = system.generate_visualization(st.session_state.strategy)
        if fig:
            st.pyplot(fig)

        # Export section
        st.subheader("5. 匯出區塊")
        if st.button("📥 下載Excel檔案", type="secondary"):
            excel_data, filename = system.export_to_excel()
            if excel_data:
                st.download_button(
                    label="點擊下載",
                    data=excel_data,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("匯出失敗")

if __name__ == "__main__":
    # Initialize session state
    if 'analysis_complete' not in st.session_state:
        st.session_state.analysis_complete = False
        st.session_state.strategy = None

    main()