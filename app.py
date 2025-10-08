"""
零售庫存調貨建議生成系統
Retail Inventory Transfer Suggestion System

版本: v1.0
開發者: Ricky
更新日期: 2025-10-08

功能概述:
- 支援三種轉貨模式：保守轉貨、加強轉貨、特強轉貨
- 智慧轉貨建議生成，基於庫存狀況和銷售數據
- 完整的統計分析和視覺化展示
- Excel匯出功能，支援企業應用

技術棧:
- 前端：Streamlit
- 資料處理：pandas, numpy
- 視覺化：matplotlib, seaborn
- Excel處理：openpyxl
"""

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import io
import xlsxwriter
import warnings
import logging

# 設定警告和日誌
warnings.filterwarnings('ignore')

# 設定頁面配置
st.set_page_config(
   page_title="調貨建議生成系統",
   page_icon="📦",
   layout="wide",
   initial_sidebar_state="expanded"
)

# 定義必要欄位
REQUIRED_COLUMNS = [
   'Article', 'Article Description', 'RP Type', 'Site', 'OM', 'MOQ',
   'SaSa Net Stock', 'Target', 'Pending Received', 'Safety Stock',
   'Last Month Sold Qty', 'MTD Sold Qty'
]

# ============================================================================
# 資料驗證與預處理函數
# ============================================================================

def validate_file_structure(df: pd.DataFrame) -> tuple[bool, str]:
   """
   驗證上傳檔案是否包含所有必要欄位

   Args:
       df: 輸入的資料框

   Returns:
       tuple: (是否有效, 訊息)
   """
   missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
   if missing_columns:
       return False, f"缺少必要欄位: {', '.join(missing_columns)}"
   return True, "檔案結構驗證通過"

def preprocess_data(df: pd.DataFrame) -> pd.DataFrame:
   """
   根據業務規則預處理資料

   Args:
       df: 原始資料框

   Returns:
       pd.DataFrame: 處理後的資料框
   """
   # 創建副本避免修改原資料
   processed_df = df.copy()

   # 初始化Notes欄位用於記錄資料清理日誌
   processed_df['Notes'] = ''

   # 1. Article欄位強制轉換為字串類型
   processed_df['Article'] = processed_df['Article'].astype(str)

   # 2. 所有數量欄位轉換為整數，無效值填充為0
   numeric_columns = [
       'MOQ', 'SaSa Net Stock', 'Target', 'Pending Received',
       'Safety Stock', 'Last Month Sold Qty', 'MTD Sold Qty'
   ]

   for col in numeric_columns:
       # 轉換為數值，錯誤值設為NaN，然後填充為0，最後轉換為整數
       processed_df[col] = pd.to_numeric(processed_df[col], errors='coerce').fillna(0).astype(int)

   # 3. 負值庫存和銷量自動修正為0
   processed_df['SaSa Net Stock'] = processed_df['SaSa Net Stock'].clip(lower=0)
   processed_df['Last Month Sold Qty'] = processed_df['Last Month Sold Qty'].clip(lower=0)
   processed_df['MTD Sold Qty'] = processed_df['MTD Sold Qty'].clip(lower=0)

   # 4. 銷量異常值（>100000）限制為100000並添加備註
   sales_outlier_mask = (processed_df['Last Month Sold Qty'] > 100000) | (processed_df['MTD Sold Qty'] > 100000)
   processed_df.loc[sales_outlier_mask, 'Notes'] = '銷量異常值已限制為100000'
   processed_df['Last Month Sold Qty'] = processed_df['Last Month Sold Qty'].clip(upper=100000)
   processed_df['MTD Sold Qty'] = processed_df['MTD Sold Qty'].clip(upper=100000)

   # 5. 字串欄位空值填充為空字串
   string_columns = ['Article Description', 'RP Type', 'Site', 'OM']
   for col in string_columns:
       processed_df[col] = processed_df[col].fillna('').astype(str)

   # 6. 驗證RP Type欄位值（只能是ND或RF）
   invalid_rp_mask = ~processed_df['RP Type'].isin(['ND', 'RF'])
   processed_df.loc[invalid_rp_mask, 'Notes'] += ' RP Type無效，已設為ND'
   processed_df.loc[invalid_rp_mask, 'RP Type'] = 'ND'

   return processed_df

def calculate_effective_sales(row: pd.Series) -> int:
   """
   根據業務規則計算有效銷量

   Args:
       row: 資料行

   Returns:
       int: 有效銷量
   """
   if row['Last Month Sold Qty'] > 0:
       return row['Last Month Sold Qty']
   else:
       return row['MTD Sold Qty']

# ============================================================================
# 轉貨候選識別函數
# ============================================================================

def identify_transfer_out_candidates(df: pd.DataFrame, mode: str = 'conservative') -> pd.DataFrame:
   """
   根據選擇的模式識別轉出候選

   Args:
       df: 處理後的資料框
       mode: 轉貨模式 ('conservative', 'enhanced', 'special')

   Returns:
       pd.DataFrame: 轉出候選清單
   """
   transfer_out_candidates = []

   # 首先計算每個產品在所有OM中的總需求
   total_demand_by_article = df[df['Target'] > 0].groupby('Article')['Target'].sum()

   # 按產品和OM分組處理
   grouped = df.groupby(['Article', 'OM'])

   for (article, om), group in grouped:
       # 為每個店鋪計算有效銷量
       group['Effective Sales'] = group.apply(calculate_effective_sales, axis=1)

       # 計算此產品在此OM中的最高銷量
       max_sales = group['Effective Sales'].max()

       # 獲取此產品在所有OM中的總需求
       article_total_demand = total_demand_by_article.get(article, 0)

       # 優先順序1：ND類型產品完全轉出
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

       # 優先順序2：RF類型產品轉出（不同模式不同邏輯）
       rf_stores = group[group['RP Type'] == 'RF']

       # 按有效銷量排序（最低的優先轉出）
       rf_stores_sorted = rf_stores.sort_values('Effective Sales', ascending=True)

       for _, store in rf_stores_sorted.iterrows():
           total_available = store['SaSa Net Stock'] + store['Pending Received']

           if mode == 'conservative':
               # 保守轉貨模式
               if total_available > store['Safety Stock']:
                   base_transferable = total_available - store['Safety Stock']
                   max_transferable = int(total_available * 0.5)  # 50%限制
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
               # 加強轉貨模式
               if total_available > (store['MOQ'] + 1):
                   base_transferable = total_available - (store['MOQ'] + 1)
                   max_transferable = int(total_available * 0.8)  # 80%限制
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
               # 特強轉貨模式
               if total_available > 0 and store['Effective Sales'] < max_sales:
                   # 基礎可轉出 = 庫存 - 2件（保留2件庫存）
                   base_transferable = store['SaSa Net Stock'] - 2
                   max_transferable = int(total_available * 0.9)  # 90%限制
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

def identify_transfer_in_candidates(df: pd.DataFrame) -> pd.DataFrame:
   """
   識別轉入候選（有目標需求的店鋪）

   Args:
       df: 處理後的資料框

   Returns:
       pd.DataFrame: 轉入候選清單
   """
   transfer_in_candidates = []

   # 篩選有目標需求的店鋪
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

# ============================================================================
# 錯誤處理函數
# ============================================================================

def handle_no_transfer_candidates(transfer_out_df: pd.DataFrame,
                               transfer_in_df: pd.DataFrame,
                               mode: str) -> dict:
   """
   處理沒有找到合格轉貨候選的情況

   Args:
       transfer_out_df: 轉出候選資料框
       transfer_in_df: 轉入候選資料框
       mode: 轉貨模式

   Returns:
       dict: 錯誤資訊和建議
   """
   # 分析情況
   no_out_candidates = transfer_out_df.empty
   no_in_candidates = transfer_in_df.empty

   # 創建診斷資訊
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
       # 檢查轉出和轉入候選的產品是否有重疊
       out_articles = set(transfer_out_df['Article'].unique())
       in_articles = set(transfer_in_df['Article'].unique())
       common_articles = out_articles.intersection(in_articles)

       if not common_articles:
           diagnostic_info['reason'] = 'no_common_articles'
           message = "沒有找到可以匹配的產品。轉出候選和轉入候選的產品沒有重疊。"
       else:
           diagnostic_info['reason'] = 'om_constraint_violation'
           message = "沒有找到符合OM約束的轉貨機會。系統要求轉出和轉入必須在同一OM單位內。"

   # 創建使用者友好的錯誤回應
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

# ============================================================================
# 轉貨匹配演算法
# ============================================================================

def match_transfers(transfer_out_df: pd.DataFrame,
                  transfer_in_df: pd.DataFrame,
                  original_df: pd.DataFrame) -> pd.DataFrame:
   """
   匹配轉出和轉入候選，生成轉貨建議

   關鍵約束：總轉出數量不能超過總需求數量

   Args:
       transfer_out_df: 轉出候選資料框
       transfer_in_df: 轉入候選資料框
       original_df: 原始資料框

   Returns:
       pd.DataFrame: 轉貨建議清單
   """
   transfer_suggestions = []

   # 檢查資料框是否為空
   if transfer_out_df.empty or transfer_in_df.empty:
       return pd.DataFrame(transfer_suggestions)

   # 創建轉入資料副本避免修改原資料
   transfer_in_df_copy = transfer_in_df.copy()

   # 按產品分組以應用約束條件
   out_grouped = transfer_out_df.groupby(['Article'])
   in_grouped = transfer_in_df_copy.groupby(['Article'])

   for article, out_group in out_grouped:
       if article in in_grouped.groups:
           in_group = in_grouped.get_group(article)

           # 計算此產品在所有OM中的總需求
           total_demand = in_group['Required Qty'].sum()

           # 獲取此產品的所有轉出候選（跨所有OM）
           out_group_sorted = out_group.sort_values(['OM', 'Transfer Type', 'Effective Sales'],
                                                  ascending=[True, True, True])

           # 轉入候選按OM和有效銷量排序（銷量高的優先）
           in_group_sorted = in_group.sort_values(['OM', 'Effective Sales'], ascending=[True, False])

           # 追蹤此產品在所有OM中的總轉出數量
           total_transferred = 0

           # 執行轉貨匹配
           for _, out_store in out_group_sorted.iterrows():
               remaining_qty = out_store['Transfer Qty']

               for idx, in_store in in_group_sorted.iterrows():
                   if remaining_qty <= 0:
                       break

                   # 避免同一店鋪自我轉貨
                   if out_store['Site'] == in_store['Site']:
                       continue

                   # 計算潛在轉移數量
                   potential_transfer_qty = min(remaining_qty, in_store['Required Qty'])

                   # 應用全域需求約束（針對此產品）
                   if total_transferred + potential_transfer_qty > total_demand:
                       potential_transfer_qty = max(0, total_demand - total_transferred)

                   if potential_transfer_qty > 0:
                       # 從原資料獲取產品描述
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

                       # 更新追蹤變數
                       remaining_qty -= potential_transfer_qty
                       total_transferred += potential_transfer_qty

                       # 更新接收店鋪的剩餘需求量（在副本中）
                       transfer_in_df_copy.loc[idx, 'Required Qty'] -= potential_transfer_qty

                       # 更新排序後的轉入群組以供下次迭代使用
                       in_group_sorted.loc[idx, 'Required Qty'] -= potential_transfer_qty

   return pd.DataFrame(transfer_suggestions)

# ============================================================================
# 統計分析函數
# ============================================================================

def calculate_statistics(transfer_suggestions_df: pd.DataFrame, mode: str) -> dict:
   """
   計算完整的統計分析，包括約束驗證

   Args:
       transfer_suggestions_df: 轉貨建議資料框
       mode: 轉貨模式

   Returns:
       dict: 統計結果
   """
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

   # 基本KPI指標
   total_transfer_qty = transfer_suggestions_df['Transfer Qty'].sum()
   total_transfer_lines = len(transfer_suggestions_df)
   unique_articles = transfer_suggestions_df['Article'].nunique()
   unique_oms = transfer_suggestions_df['OM'].nunique()

   # 計算每個產品的總需求和總轉出（用於約束驗證）
   total_demand_by_article = transfer_suggestions_df.groupby('Article')['Receive Site Target Qty'].first()
   total_transfer_by_article = transfer_suggestions_df.groupby('Article')['Transfer Qty'].sum()

   # 檢查約束違規
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

   # 按產品統計
   article_stats = transfer_suggestions_df.groupby('Article').agg({
       'Receive Site Target Qty': 'first',  # 總需求件數
       'Transfer Qty': 'sum',  # 總調貨件數
       'OM': 'nunique'  # 涉及OM數量
   }).round(2)
   article_stats.columns = ['總需求件數', '總調貨件數', '涉及OM數量']
   article_stats['轉貨行數'] = transfer_suggestions_df.groupby('Article').size()
   article_stats['需求滿足率'] = (article_stats['總調貨件數'] / article_stats['總需求件數'] * 100).round(2)
   article_stats['約束違規'] = [(total_transfer_by_article.get(article, 0) > article_stats.loc[article, '總需求件數']) for article in article_stats.index]

   # 按OM統計
   om_stats = transfer_suggestions_df.groupby('OM').agg({
       'Receive Site Target Qty': 'first',  # 總需求件數
       'Transfer Qty': 'sum',  # 總調貨件數
       'Article': 'nunique'  # 涉及產品數量
   }).round(2)
   om_stats.columns = ['總需求件數', '總調貨件數', '涉及產品數量']
   om_stats['轉貨行數'] = transfer_suggestions_df.groupby('OM').size()

   # 轉出類型分佈
   transfer_type_stats = transfer_suggestions_df.groupby('Transfer Type').agg({
       'Transfer Qty': ['sum', 'count']
   }).round(2)
   transfer_type_stats.columns = ['總件數', '涉及行數']

   # 接收統計
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

# ============================================================================
# 視覺化函數
# ============================================================================

def create_visualization(stats: dict, transfer_suggestions_df: pd.DataFrame, mode: str):
   """
   根據模式創建matplotlib視覺化圖表

   Args:
       stats: 統計資料
       transfer_suggestions_df: 轉貨建議資料框
       mode: 轉貨模式

   Returns:
       matplotlib圖表物件或None
   """
   if transfer_suggestions_df.empty:
       return None

   fig, ax = plt.subplots(figsize=(14, 8))

   # 準備轉出資料（按OM和轉貨類型）
   transfer_out_by_om_type = transfer_suggestions_df.groupby(['OM', 'Transfer Type'])['Transfer Qty'].sum().unstack(fill_value=0)

   # 準備接收資料（按OM）
   receive_data = transfer_suggestions_df.groupby('Receive Site')['Transfer Qty'].sum()
   receive_by_om = transfer_suggestions_df.drop_duplicates('Receive Site').set_index('Receive Site')['OM']
   receive_by_om = receive_by_om[receive_data.index]
   receive_by_om_grouped = receive_by_om.groupby(receive_by_om).sum().rename('實際接收數量')

   # 準備目標資料（按OM）
   target_by_om = transfer_suggestions_df.drop_duplicates('Receive Site').groupby('OM')['Receive Site Target Qty'].sum().rename('需求接收數量')

   # 合併所有資料
   combined_data = transfer_out_by_om_type.join(receive_by_om_grouped).join(target_by_om).fillna(0)

   # 根據模式定義預期的欄位
   if mode == 'conservative':
       # 模式A：4條形設計
       expected_columns = ['ND轉出', 'RF過剩轉出', '需求接收數量', '實際接收數量']
   elif mode == 'enhanced':
       # 模式B：5條形設計
       expected_columns = ['ND轉出', 'RF過剩轉出', 'RF加強轉出', '需求接收數量', '實際接收數量']
   else:  # special mode
       # 模式C：5條形設計
       expected_columns = ['ND轉出', 'RF過剩轉出', 'RF特強轉出', '需求接收數量', '實際接收數量']

   # 篩選並重新排序欄位
   available_columns = [col for col in expected_columns if col in combined_data.columns]
   combined_data = combined_data[available_columns]

   # 創建長條圖
   combined_data.plot(kind='bar', ax=ax, width=0.8)

   ax.set_title('Transfer Receive Analysis', fontsize=16, fontweight='bold')
   ax.set_xlabel('OM單位', fontsize=12)
   ax.set_ylabel('調貨數量', fontsize=12)
   ax.legend(title='轉貨類型', bbox_to_anchor=(1.05, 1), loc='upper left')
   ax.grid(axis='y', alpha=0.3)

   plt.xticks(rotation=45)
   plt.tight_layout()

   return fig

# ============================================================================
# Excel匯出函數
# ============================================================================

def export_to_excel(transfer_suggestions_df: pd.DataFrame, stats: dict) -> io.BytesIO:
   """
   匯出結果到Excel檔案，格式完全符合需求

   Args:
       transfer_suggestions_df: 轉貨建議資料框
       stats: 統計資料

   Returns:
       io.BytesIO: Excel檔案內容
   """
   output = io.BytesIO()

   with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
       # 工作表1：轉貨建議（特定欄位順序）
       required_columns = [
           'Article', 'Product Desc', 'OM', 'Transfer Site', 'Transfer Qty',
           'Transfer Site Original Stock', 'Transfer Site After Transfer Stock',
           'Transfer Site Safety Stock', 'Transfer Site MOQ', 'Receive Site',
           'Receive Site Target Qty', 'Notes'
       ]

       # 篩選存在的欄位
       available_columns = [col for col in required_columns if col in transfer_suggestions_df.columns]
       export_df = transfer_suggestions_df[available_columns].copy()

       # 重新命名欄位為英文（系統相容性）
       column_rename_map = {
           'Article': 'Article',
           'Product Desc': 'Product Description',
           'OM': 'OM',
           'Transfer Site': 'Transfer Site',
           'Transfer Qty': 'Transfer Qty',
           'Transfer Site Original Stock': 'Transfer Site Original Stock',
           'Transfer Site After Transfer Stock': 'Transfer Site After Transfer Stock',
           'Transfer Site Safety Stock': 'Transfer Site Safety Stock',
           'Transfer Site MOQ': 'Transfer Site MOQ',
           'Receive Site': 'Receive Site',
           'Receive Site Target Qty': 'Receive Site Target Qty',
           'Notes': 'Notes'
       }

       export_df = export_df.rename(columns=column_rename_map)
       export_df.to_excel(writer, sheet_name='調貨建議', index=False)

       # 工作表2：統計摘要（適當間距）
       workbook = writer.book
       worksheet = workbook.add_worksheet('統計摘要')

       # KPI概覽
       row = 0
       worksheet.write(row, 0, 'KPI Overview')
       worksheet.write(row, 1, '總轉貨建議數量')
       worksheet.write(row, 2, stats['total_transfer_qty'])
       row += 1
       worksheet.write(row, 1, '總轉貨件數')
       worksheet.write(row, 2, stats['total_transfer_lines'])
       row += 1
       worksheet.write(row, 1, '涉及產品數量')
       worksheet.write(row, 2, stats['unique_articles'])
       row += 1
       worksheet.write(row, 1, '涉及OM數量')
       worksheet.write(row, 2, stats['unique_oms'])
       row += 1

       # 留3行空白
       row += 3

       # 按產品統計
       if not stats['article_stats'].empty:
           worksheet.write(row, 0, 'Statistics by Article')
           row += 1
           # 寫入標題
           headers = ['Article', '總需求件數', '總調貨件數', '涉及OM數量', '轉貨行數', '需求滿足率', '約束違規']
           for col_num, header in enumerate(headers):
               worksheet.write(row, col_num, header)
           row += 1

           # 寫入資料
           for article in stats['article_stats'].index:
               base_col = 0
               worksheet.write(row, base_col, article)
               worksheet.write(row, base_col + 1, stats['article_stats'].loc[article, '總需求件數'])
               worksheet.write(row, base_col + 2, stats['article_stats'].loc[article, '總調貨件數'])
               worksheet.write(row, base_col + 3, stats['article_stats'].loc[article, '涉及OM數量'])
               worksheet.write(row, base_col + 4, stats['article_stats'].loc[article, '轉貨行數'])
               worksheet.write(row, base_col + 5, stats['article_stats'].loc[article, '需求滿足率'])
               worksheet.write(row, base_col + 6, '是' if stats['article_stats'].loc[article, '約束違規'] else '否')
               row += 1

           row += 3  # 留3行空白

       # 按OM統計
       if not stats['om_stats'].empty:
           worksheet.write(row, 0, 'Statistics by OM')
           row += 1
           # 寫入標題
           headers = ['OM', '總需求件數', '總調貨件數', '涉及產品數量', '轉貨行數']
           for col_num, header in enumerate(headers):
               worksheet.write(row, col_num, header)
           row += 1

           # 寫入資料
           for om in stats['om_stats'].index:
               base_col = 0
               worksheet.write(row, base_col, om)
               worksheet.write(row, base_col + 1, stats['om_stats'].loc[om, '總需求件數'])
               worksheet.write(row, base_col + 2, stats['om_stats'].loc[om, '總調貨件數'])
               worksheet.write(row, base_col + 3, stats['om_stats'].loc[om, '涉及產品數量'])
               worksheet.write(row, base_col + 4, stats['om_stats'].loc[om, '轉貨行數'])
               row += 1

           row += 3  # 留3行空白

       # 轉出類型分佈
       if not stats['transfer_type_stats'].empty:
           worksheet.write(row, 0, 'Transfer Type Distribution')
           row += 1
           # 寫入標題
           headers = ['Transfer Type', '總件數', '涉及行數']
           for col_num, header in enumerate(headers):
               worksheet.write(row, col_num, header)
           row += 1

           # 寫入資料
           for transfer_type in stats['transfer_type_stats'].index:
               base_col = 0
               worksheet.write(row, base_col, transfer_type)
               worksheet.write(row, base_col + 1, stats['transfer_type_stats'].loc[transfer_type, '總件數'])
               worksheet.write(row, base_col + 2, stats['transfer_type_stats'].loc[transfer_type, '涉及行數'])
               row += 1

           row += 3  # 留3行空白

       # 接收類型分佈
       if not stats['receive_stats'].empty:
           worksheet.write(row, 0, 'Receive Type Distribution')
           row += 1
           # 寫入標題
           headers = ['Receive Site', '實際接收數量', '目標需求數量', '需求滿足率']
           for col_num, header in enumerate(headers):
               worksheet.write(row, col_num, header)
           row += 1

           # 寫入資料
           for receive_site in stats['receive_stats'].index:
               base_col = 0
               worksheet.write(row, base_col, receive_site)
               worksheet.write(row, base_col + 1, stats['receive_stats'].loc[receive_site, '實際接收數量'])
               worksheet.write(row, base_col + 2, stats['receive_stats'].loc[receive_site, '目標需求數量'])
               worksheet.write(row, base_col + 3, stats['receive_stats'].loc[receive_site, '需求滿足率'])
               row += 1

   output.seek(0)
   return output

# ============================================================================
# 主應用程式
# ============================================================================

def main():
   """
   Streamlit主應用程式
   """
   # 頁面標頭
   st.title("📦 調貨建議生成系統")
   st.markdown("---")

   # 側邊欄
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

   # 檔案上傳區塊
   st.header("1. 資料上傳")
   uploaded_file = st.file_uploader(
       "請上傳Excel檔案 (.xlsx, .xls)",
       type=['xlsx', 'xls'],
       help="檔案必須包含所有必要的欄位：Article, Article Description, RP Type, Site, OM, MOQ, SaSa Net Stock, Target, Pending Received, Safety Stock, Last Month Sold Qty, MTD Sold Qty"
   )

   if uploaded_file is not None:
       try:
           # 讀取上傳的檔案
           df = pd.read_excel(uploaded_file)

           # 驗證檔案結構
           is_valid, message = validate_file_structure(df)

           if not is_valid:
               st.error(f"❌ {message}")
               return

           # 預處理資料
           with st.spinner("正在處理資料..."):
               df = preprocess_data(df)

           st.success("✅ 檔案上傳並處理成功！")

           # 資料預覽區塊
           st.header("2. 資料預覽")

           col1, col2, col3 = st.columns(3)
           with col1:
               st.metric("總記錄數", len(df))
           with col2:
               st.metric("產品數量", df['Article'].nunique())
           with col3:
               st.metric("店鋪數量", df['Site'].nunique())

           # 顯示樣本資料
           st.subheader("資料樣本")
           st.dataframe(df.head(10))

           # 轉貨模式選擇
           st.header("3. 轉貨模式選擇")
           mode = st.radio(
               "請選擇轉貨模式：",
               options=["conservative", "enhanced", "special"],
               format_func=lambda x: "A: 保守轉貨" if x == "conservative" else ("B: 加強轉貨" if x == "enhanced" else "C: 特強轉貨"),
               help="保守轉貨：RF類型轉出限制為50% | 加強轉貨：RF類型轉出限制為80% | 特強轉貨：RF類型轉出限制為90%並保留2件庫存"
           )

           # 分析執行區塊
           st.header("4. 分析執行")
           if st.button("🚀 開始分析", type="primary", use_container_width=True):
               with st.spinner("正在生成轉貨建議..."):
                   # 識別轉貨候選
                   transfer_out_df = identify_transfer_out_candidates(df, mode)
                   transfer_in_df = identify_transfer_in_candidates(df)

                   # 匹配轉貨
                   transfer_suggestions_df = match_transfers(transfer_out_df, transfer_in_df, df)

                   # 計算統計
                   stats = calculate_statistics(transfer_suggestions_df, mode)

               st.success("✅ 分析完成！")

               # 結果展示區塊
               st.header("5. 分析結果")

               # 檢查約束違規
               if stats.get('constraint_violations', 0) > 0:
                   st.error(f"⚠️ 發現 {stats['constraint_violations']} 個約束違規：總轉出數量超過總需求數量")

                   # 在可展開區塊顯示違規詳情
                   with st.expander("約束違規詳情"):
                       for violation in stats.get('violation_details', []):
                           st.write(f"**產品 {violation['Article']}**:")
                           st.write(f"  - 總需求: {violation['Total Demand']}")
                           st.write(f"  - 總轉出: {violation['Total Transfer']}")
                           st.write(f"  - 違規數量: {violation['Violation']}")
               else:
                   # 添加約束合規指示器
                   if stats.get('total_transfer_qty', 0) > 0:
                       st.success("✅ 所有轉貨建議均符合需求約束")

               # KPI指標
               col1, col2, col3, col4 = st.columns(4)
               with col1:
                   st.metric("總轉貨建議數量", stats['total_transfer_qty'])
               with col2:
                   st.metric("總轉貨件數", stats['total_transfer_lines'])
               with col3:
                   st.metric("涉及產品數量", stats['unique_articles'])
               with col4:
                   st.metric("涉及OM數量", stats['unique_oms'])

               # 轉貨建議表格
               st.subheader("轉貨建議明細")
               if not transfer_suggestions_df.empty:
                   st.dataframe(transfer_suggestions_df, use_container_width=True)
               else:
                   # 使用錯誤處理函數
                   error_info = handle_no_transfer_candidates(transfer_out_df, transfer_in_df, mode)

                   # 顯示使用者友好的訊息
                   st.warning(f"⚠️ {error_info['message']}")

                   # 在可展開區塊顯示建議
                   with st.expander("疑難排解建議"):
                       st.write("**建議解決方案：**")
                       for suggestion in error_info['suggestions']:
                           st.write(f"• {suggestion}")

               # 統計分析表格
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

                   # 數據視覺化
                   st.subheader("數據視覺化")
                   fig = create_visualization(stats, transfer_suggestions_df, mode)
                   if fig:
                       st.pyplot(fig)
                   else:
                       st.info("沒有足夠的數據生成圖表")
               else:
                   st.info("📊 沒有轉貨建議資料，無法生成統計分析和圖表")

               # 匯出區塊
               st.header("6. 匯出功能")

               if not transfer_suggestions_df.empty:
                   # 生成Excel檔案
                   excel_data = export_to_excel(transfer_suggestions_df, stats)

                   # 創建下載按鈕
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