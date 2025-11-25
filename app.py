import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage
import io
import tempfile
import os
import re
from collections import defaultdict

# Set page configuration
st.set_page_config(
    page_title="Sales Report Generator",
    page_icon="📊",
    layout="wide"
)

# Title and description
st.title("📊 Sales Report Generator")
st.markdown("""
Upload your sales Excel file to generate a comprehensive sales summary report with visualizations.
The app will calculate MTD and YTD metrics, compare year-over-year performance, and create professional charts.
""")

# Sidebar for configuration
st.sidebar.header("⚙️ Configuration")
target_month = st.sidebar.slider("Select Target Month", 1, 12, 10, help="Month for MTD calculations")
target_year = st.sidebar.number_input("Select Target Year", 2020, 2030, 2025, help="Year for calculations")
comparison_year = st.sidebar.number_input("Comparison Year", 2020, 2030, 2024, help="Previous year for comparison")

# Month names for display
month_names = ["", "January", "February", "March", "April", "May", "June", 
               "July", "August", "September", "October", "November", "December"]
month_name = month_names[target_month]

# File uploader
uploaded_file = st.file_uploader("Upload Sales Excel File (.xls or .xlsx)", type=['xls', 'xlsx'])

def split_by_items(df):
    """Split dataframe by item numbers"""
    items_dict = {}
    current_item = None
    current_data = []
    column_names = df.columns

    def finalize_item(rows, item_key):
        if not rows:
            return None
        item_df = pd.DataFrame(rows, columns=column_names)
        item_df = item_df.dropna(axis=1, how='all')
        if 'Item Number' in item_df.columns:
            item_df['Item Number'] = item_df['Item Number'].ffill().bfill().fillna(item_key)
        else:
            item_df.insert(0, 'Item Number', item_key)
        return item_df
    
    for idx, row in df.iterrows():
        first_col_value = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ''
        first_col_clean = first_col_value.strip()
        
        is_item_code = (first_col_clean and 
                       any(c.isalpha() for c in first_col_clean) and 
                       any(c.isdigit() for c in first_col_clean) and
                       'Item Total' not in first_col_value)
        
        if is_item_code:
            if current_item and current_data:
                item_df = finalize_item(current_data, current_item)
                if item_df is not None:
                    items_dict[current_item] = item_df
            
            current_item = first_col_clean
            current_data = [row.values]
        elif current_item:
            current_data.append(row.values)
            
            if 'Item Total:' in str(row.values):
                item_df = finalize_item(current_data, current_item)
                if item_df is not None:
                    items_dict[current_item] = item_df
                current_item = None
                current_data = []
    
    if current_item and current_data:
        item_df = finalize_item(current_data, current_item)
        if item_df is not None:
            items_dict[current_item] = item_df
    
    return items_dict

def calculate_sales_metrics(all_items, target_month, target_year):
    """Calculate MTD and YTD sales metrics"""
    mtd_sales = 0
    mtd_cost = 0
    ytd_sales = 0
    ytd_cost = 0
    items_with_data = 0
    
    for item_key, item_df in all_items.items():
        if item_df.empty:
            continue
        
        df = item_df.copy()
        
        # Find column names
        year_col = None
        period_col = None
        sales_col = None
        cost_col = None
        
        for col in df.columns:
            col_lower = str(col).strip().lower()
            if 'year' in col_lower and not year_col:
                year_col = col
            elif 'period' in col_lower and not period_col:
                period_col = col
            elif 'sales amount' in col_lower and not sales_col:
                sales_col = col
            elif 'cost of sales' in col_lower and not cost_col:
                cost_col = col
        
        if not year_col or not period_col or not sales_col:
            continue
        
        # Convert to numeric
        df[year_col] = pd.to_numeric(df[year_col], errors='coerce')
        df[period_col] = pd.to_numeric(df[period_col], errors='coerce')
        df[sales_col] = pd.to_numeric(df[sales_col], errors='coerce').fillna(0)
        if cost_col:
            df[cost_col] = pd.to_numeric(df[cost_col], errors='coerce').fillna(0)
        
        year_data = df[df[year_col] == target_year].copy()
        
        if not year_data.empty:
            items_with_data += 1
            
            # MTD
            mtd_data = year_data[year_data[period_col] == target_month]
            mtd_sales += mtd_data[sales_col].sum()
            if cost_col:
                mtd_cost += mtd_data[cost_col].sum()
            
            # YTD
            ytd_data = year_data[year_data[period_col] <= target_month]
            ytd_sales += ytd_data[sales_col].sum()
            if cost_col:
                ytd_cost += ytd_data[cost_col].sum()
    
    mtd_gp = ((mtd_sales - mtd_cost) / mtd_sales * 100) if mtd_sales > 0 else 0
    ytd_gp = ((ytd_sales - ytd_cost) / ytd_sales * 100) if ytd_sales > 0 else 0
    
    return {
        'MTD Gross Sales': mtd_sales,
        'MTD GP%': mtd_gp,
        'YTD Gross Sales': ytd_sales,
        'YTD GP%': ytd_gp,
        'items_processed': items_with_data
    }

def create_excel_report(results_current, results_previous, target_month, target_year, comparison_year, month_name):
    """Create Excel report with formatted table and charts"""
    
    # Calculate % achieved
    mtd_achieved = (results_current['MTD Gross Sales'] / results_previous['MTD Gross Sales'] * 100) if results_previous['MTD Gross Sales'] > 0 else 0
    ytd_achieved = (results_current['YTD Gross Sales'] / results_previous['YTD Gross Sales'] * 100) if results_previous['YTD Gross Sales'] > 0 else 0
    mtd_gp_achieved = "0%" if results_previous['MTD GP%'] == 0 else f"{(results_current['MTD GP%'] / results_previous['MTD GP%'] * 100):.0f}%"
    ytd_gp_achieved = "0%" if results_previous['YTD GP%'] == 0 else f"{(results_current['YTD GP%'] / results_previous['YTD GP%'] * 100):.0f}%"
    
    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales Summary Report"
    
    # Title
    ws['A1'] = f"Sales Summary Report - {month_name} {target_year}"
    ws['A1'].font = Font(size=14, bold=True)
    ws['A1'].alignment = Alignment(horizontal='center')
    ws.merge_cells('A1:E1')
    
    # Styling
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    row_current_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    row_achieved_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    row_budget_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    
    border_style = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Headers
    headers = ['Date', 'MTD Gross Sales', 'MTD GP%', 'YTD Gross Sales', 'YTD GP%']
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border_style
    
    # Data rows
    data_rows = [
        [str(target_year), 
         results_current['MTD Gross Sales'], 
         results_current['MTD GP%'] / 100,
         results_current['YTD Gross Sales'], 
         results_current['YTD GP%'] / 100],
        [str(comparison_year), 
         results_previous['MTD Gross Sales'], 
         results_previous['MTD GP%'] / 100, 
         results_previous['YTD Gross Sales'], 
         results_previous['YTD GP%'] / 100],
        ['%Achieved', 
         mtd_achieved / 100 if results_previous['MTD Gross Sales'] > 0 else 0,
         float(mtd_gp_achieved.rstrip('%')) / 100 if mtd_gp_achieved != "0%" else 0,
         ytd_achieved / 100 if results_previous['YTD Gross Sales'] > 0 else 0,
         float(ytd_gp_achieved.rstrip('%')) / 100 if ytd_gp_achieved != "0%" else 0],
        [f'{target_year} Budget', '', '', '', ''],
        ['% Achieved', 0, 0, 0, 0]
    ]
    
    # Write data rows
    for row_idx, row_data in enumerate(data_rows, start=4):
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border_style
            cell.alignment = Alignment(horizontal='right' if col_idx > 1 else 'left', vertical='center')
            
            if row_data[0] == str(target_year):
                cell.fill = row_current_fill
            elif '%Achieved' in row_data[0]:
                cell.fill = row_achieved_fill
            elif 'Budget' in row_data[0]:
                cell.fill = row_budget_fill
            
            if col_idx > 1 and value != '':
                if col_idx in [2, 4]:
                    cell.number_format = '$#,##0.00'
                elif col_idx in [3, 5]:
                    cell.number_format = '0.00%'
    
    # Set column widths
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 12
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[3].height = 20
    
    # Create dashboard chart
    fig_dashboard, axes = plt.subplots(2, 2, figsize=(12, 9))
    fig_dashboard.suptitle(f'Sales Performance Dashboard - {month_name} {target_year}', fontsize=14, fontweight='bold')
    
    categories = [str(target_year), str(comparison_year)]
    mtd_sales_values = [results_current['MTD Gross Sales'], results_previous['MTD Gross Sales']]
    ytd_sales_values = [results_current['YTD Gross Sales'], results_previous['YTD Gross Sales']]
    mtd_gp_values = [results_current['MTD GP%'], results_previous['MTD GP%']]
    ytd_gp_values = [results_current['YTD GP%'], results_previous['YTD GP%']]
    colors = ['#2E86AB', '#A23B72']
    
    # MTD Gross Sales
    ax1 = axes[0, 0]
    bars1 = ax1.bar(categories, mtd_sales_values, color=colors, alpha=0.8, edgecolor='black', linewidth=1.5)
    ax1.set_title('MTD Gross Sales', fontsize=11, fontweight='bold')
    ax1.set_ylabel('Sales Amount ($)', fontsize=9)
    ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
    for bar in bars1:
        height = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2., height, f'${height:,.0f}',
                ha='center', va='bottom', fontsize=8, fontweight='bold')
    if results_previous['MTD Gross Sales'] > 0:
        ax1.text(0.5, max(mtd_sales_values) * 0.92, f'% Achieved: {mtd_achieved:.0f}%',
                ha='center', fontsize=9, bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
    ax1.grid(axis='y', alpha=0.3, linestyle='--')
    
    # YTD Gross Sales
    ax2 = axes[0, 1]
    bars2 = ax2.bar(categories, ytd_sales_values, color=colors, alpha=0.8, edgecolor='black', linewidth=1.5)
    ax2.set_title('YTD Gross Sales', fontsize=11, fontweight='bold')
    ax2.set_ylabel('Sales Amount ($)', fontsize=9)
    ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
    for bar in bars2:
        height = bar.get_height()
        ax2.text(bar.get_x() + bar.get_width()/2., height, f'${height:,.0f}',
                ha='center', va='bottom', fontsize=8, fontweight='bold')
    if results_previous['YTD Gross Sales'] > 0:
        ax2.text(0.5, max(ytd_sales_values) * 0.92, f'% Achieved: {ytd_achieved:.0f}%',
                ha='center', fontsize=9, bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
    ax2.grid(axis='y', alpha=0.3, linestyle='--')
    
    # MTD GP%
    ax3 = axes[1, 0]
    bars3 = ax3.bar(categories, mtd_gp_values, color=colors, alpha=0.8, edgecolor='black', linewidth=1.5)
    ax3.set_title('MTD Gross Profit %', fontsize=11, fontweight='bold')
    ax3.set_ylabel('GP %', fontsize=9)
    ax3.set_ylim(0, max(mtd_gp_values) * 1.2 if max(mtd_gp_values) > 0 else 100)
    for bar in bars3:
        height = bar.get_height()
        ax3.text(bar.get_x() + bar.get_width()/2., height, f'{height:.2f}%',
                ha='center', va='bottom', fontsize=8, fontweight='bold')
    if results_previous['MTD GP%'] > 0:
        gp_achieved = (results_current['MTD GP%'] / results_previous['MTD GP%'] * 100)
        ax3.text(0.5, max(mtd_gp_values) * 0.92, f'% Achieved: {gp_achieved:.0f}%',
                ha='center', fontsize=9, bbox=dict(boxstyle='round', facecolor='lightgreen', alpha=0.5))
    ax3.grid(axis='y', alpha=0.3, linestyle='--')
    
    # YTD GP%
    ax4 = axes[1, 1]
    bars4 = ax4.bar(categories, ytd_gp_values, color=colors, alpha=0.8, edgecolor='black', linewidth=1.5)
    ax4.set_title('YTD Gross Profit %', fontsize=11, fontweight='bold')
    ax4.set_ylabel('GP %', fontsize=9)
    ax4.set_ylim(0, max(ytd_gp_values) * 1.2 if max(ytd_gp_values) > 0 else 100)
    for bar in bars4:
        height = bar.get_height()
        ax4.text(bar.get_x() + bar.get_width()/2., height, f'{height:.2f}%',
                ha='center', va='bottom', fontsize=8, fontweight='bold')
    if results_previous['YTD GP%'] > 0:
        gp_achieved = (results_current['YTD GP%'] / results_previous['YTD GP%'] * 100)
        ax4.text(0.5, max(ytd_gp_values) * 0.92, f'% Achieved: {gp_achieved:.0f}%',
                ha='center', fontsize=9, bbox=dict(boxstyle='round', facecolor='lightgreen', alpha=0.5))
    ax4.grid(axis='y', alpha=0.3, linestyle='--')
    
    plt.tight_layout()
    
    # Save dashboard chart
    img_buffer1 = io.BytesIO()
    fig_dashboard.savefig(img_buffer1, format='png', dpi=150, bbox_inches='tight')
    img_buffer1.seek(0)
    plt.close(fig_dashboard)
    
    ws_chart1 = wb.create_sheet("Dashboard Chart")
    img1 = XLImage(img_buffer1)
    ws_chart1.add_image(img1, 'A1')
    
    # Create comparison chart
    fig_comparison, ax = plt.subplots(1, 1, figsize=(11, 6))
    x = np.arange(4)
    width = 0.35
    metrics = ['MTD Gross Sales\n(in thousands)', 'MTD GP%', 'YTD Gross Sales\n(in thousands)', 'YTD GP%']
    values_current = [
        results_current['MTD Gross Sales'] / 1000,
        results_current['MTD GP%'],
        results_current['YTD Gross Sales'] / 1000,
        results_current['YTD GP%']
    ]
    values_previous = [
        results_previous['MTD Gross Sales'] / 1000,
        results_previous['MTD GP%'],
        results_previous['YTD Gross Sales'] / 1000,
        results_previous['YTD GP%']
    ]
    
    bars1 = ax.bar(x - width/2, values_current, width, label=str(target_year), color='#2E86AB', alpha=0.8, edgecolor='black', linewidth=1.5)
    bars2 = ax.bar(x + width/2, values_previous, width, label=str(comparison_year), color='#A23B72', alpha=0.8, edgecolor='black', linewidth=1.5)
    
    ax.set_title('Year-over-Year Performance Comparison', fontsize=13, fontweight='bold')
    ax.set_ylabel('Value', fontsize=10)
    ax.set_xticks(x)
    ax.set_xticklabels(metrics, fontsize=9)
    ax.legend(fontsize=10, loc='upper left')
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    
    for bars in [bars1, bars2]:
        for i, bar in enumerate(bars):
            height = bar.get_height()
            if i in [0, 2]:
                label = f'${height:,.0f}K'
            else:
                label = f'{height:.2f}%'
            ax.text(bar.get_x() + bar.get_width()/2., height,
                    label, ha='center', va='bottom', fontsize=7, fontweight='bold')
    
    plt.tight_layout()
    
    # Save comparison chart
    img_buffer2 = io.BytesIO()
    fig_comparison.savefig(img_buffer2, format='png', dpi=150, bbox_inches='tight')
    img_buffer2.seek(0)
    plt.close(fig_comparison)
    
    ws_chart2 = wb.create_sheet("Comparison Chart")
    img2 = XLImage(img_buffer2)
    ws_chart2.add_image(img2, 'A1')
    
    # Save to bytes
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output

def extract_brand_code(item_key):
    """Extract letter prefix from item code (e.g., AMO0002 -> AMO)"""
    code = str(item_key).split('_')[-1]
    m = re.match(r'([A-Za-z]+)', code)
    return m.group(1).upper() if m else code.upper()

def compute_item_year_metrics(item_df, year, month):
    """Compute metrics for a single item dataframe and a given year"""
    df = item_df.copy()
    year_col = None
    period_col = None
    sales_col = None
    cost_col = None
    desc_col = None
    
    for col in df.columns:
        col_lower = str(col).strip().lower()
        if 'year' in col_lower and not year_col:
            year_col = col
        elif 'period' in col_lower and not period_col:
            period_col = col
        elif 'sales amount' in col_lower and not sales_col:
            sales_col = col
        elif 'cost of sales' in col_lower and not cost_col:
            cost_col = col
        elif any(k in col_lower for k in ['description','item','name']) and not desc_col:
            desc_col = col

    if not year_col or not period_col or not sales_col:
        return {'mtd_sales':0.0,'mtd_cost':0.0,'ytd_sales':0.0,'ytd_cost':0.0,'desc':None}

    df[year_col] = pd.to_numeric(df[year_col], errors='coerce')
    df[period_col] = pd.to_numeric(df[period_col], errors='coerce')
    df[sales_col] = pd.to_numeric(df[sales_col], errors='coerce').fillna(0)
    if cost_col:
        df[cost_col] = pd.to_numeric(df[cost_col], errors='coerce').fillna(0)

    year_data = df[df[year_col] == year]
    if year_data.empty:
        return {'mtd_sales':0.0,'mtd_cost':0.0,'ytd_sales':0.0,'ytd_cost':0.0,'desc':None}

    mtd_data = year_data[year_data[period_col] == month]
    ytd_data = year_data[year_data[period_col] <= month]

    mtd_sales = float(mtd_data[sales_col].sum())
    mtd_cost = float(mtd_data[cost_col].sum()) if cost_col else 0.0
    ytd_sales = float(ytd_data[sales_col].sum())
    ytd_cost = float(ytd_data[cost_col].sum()) if cost_col else 0.0

    desc_val = None
    if desc_col is not None:
        non_nulls = df[desc_col].dropna().astype(str)
        if len(non_nulls):
            desc_val = non_nulls.iloc[0]

    return {'mtd_sales':mtd_sales,'mtd_cost':mtd_cost,'ytd_sales':ytd_sales,'ytd_cost':ytd_cost,'desc':desc_val}

def create_sku_report(all_items, target_month, target_year, comparison_year, month_name, report_type='MTD'):
    """Create Top-20 SKU MTD or YTD Performance report (excludes MX products)"""
    
    # Collect all SKUs with their metrics
    sku_metrics = []
    
    for item_key, item_df in all_items.items():
        if item_df.empty:
            continue

        sku_code, item_name = get_sku_code_and_name(item_df)
        if not sku_code:
            continue

        brand_code = extract_brand_code(sku_code)
        if brand_code == 'MX':
            continue
        
        # Compute metrics for both years
        metrics_current = compute_item_year_metrics(item_df, target_year, target_month)
        metrics_previous = compute_item_year_metrics(item_df, comparison_year, target_month)
        
        # Calculate GP%
        def compute_gp(sales, cost):
            return ((sales - cost) / sales * 100) if sales > 0 else 0.0
        
        current_mtd_gp = compute_gp(metrics_current['mtd_sales'], metrics_current['mtd_cost'])
        current_ytd_gp = compute_gp(metrics_current['ytd_sales'], metrics_current['ytd_cost'])
        previous_mtd_gp = compute_gp(metrics_previous['mtd_sales'], metrics_previous['mtd_cost'])
        previous_ytd_gp = compute_gp(metrics_previous['ytd_sales'], metrics_previous['ytd_cost'])
        
        # Store metrics
        sku_metrics.append({
            'code': sku_code,
            'name': item_name,
            f'{target_year}_mtd_sales': metrics_current['mtd_sales'],
            f'{target_year}_mtd_qty': 0,  # Quantity not available in current data structure
            f'{target_year}_mtd_gp': current_mtd_gp,
            f'{comparison_year}_mtd_sales': metrics_previous['mtd_sales'],
            f'{comparison_year}_mtd_qty': 0,
            f'{comparison_year}_mtd_gp': previous_mtd_gp,
            f'{target_year}_ytd_sales': metrics_current['ytd_sales'],
            f'{target_year}_ytd_qty': 0,
            f'{target_year}_ytd_gp': current_ytd_gp,
            f'{comparison_year}_ytd_sales': metrics_previous['ytd_sales'],
            f'{comparison_year}_ytd_qty': 0,
            f'{comparison_year}_ytd_gp': previous_ytd_gp,
        })
    
    # Convert to DataFrame
    sku_df = pd.DataFrame(sku_metrics)
    if sku_df.empty:
        return None
    
    # Sort by MTD or YTD sales and get top 20
    sort_key = f'{target_year}_{report_type.lower()}_sales'
    sku_df = sku_df.sort_values(by=sort_key, ascending=False).head(20)
    
    # Create Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.title = f'Top 20 {report_type} SKUs'
    
    # Title
    ws.merge_cells('A1:I1')
    tcell = ws['A1']
    tcell.value = f'Top 20 {report_type} SKUs Performance'
    tcell.font = Font(size=14, bold=True)
    tcell.alignment = Alignment(horizontal='center')
    
    # Header row styling
    header_fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    header_fill_green = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    
    metric_lower = report_type.lower()
    
    headers = [
        ('', None),
        ('Code', None),
        ('Item Name', None),
        (f'{target_year} {report_type}\nGross Sales', header_fill_yellow),
        (f'{target_year} {report_type}\nQTY', header_fill_yellow),
        (f'{target_year} {report_type}\nGP%', header_fill_yellow),
        (f'{comparison_year} {report_type}\nGross Sales', header_fill_green),
        (f'{comparison_year} {report_type}\nQTY', header_fill_green),
        (f'{comparison_year} {report_type} GP%', header_fill_green)
    ]
    
    for c_idx, (h, fill) in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=c_idx, value=h)
        cell.font = Font(bold=True, size=10)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                           top=Side(style='thin'), bottom=Side(style='thin'))
        if fill:
            cell.fill = fill
    
    # Write data rows
    for r_idx, row in sku_df.reset_index(drop=True).iterrows():
        rank = r_idx + 1
        row_data = [
            rank,
            row['code'],
            row['name'],
            row[f'{target_year}_{metric_lower}_sales'],
            row[f'{target_year}_{metric_lower}_qty'],
            row[f'{target_year}_{metric_lower}_gp'] / 100,
            row[f'{comparison_year}_{metric_lower}_sales'],
            row[f'{comparison_year}_{metric_lower}_qty'],
            row[f'{comparison_year}_{metric_lower}_gp'] / 100
        ]
        
        excel_row = r_idx + 3
        
        for c_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=excel_row, column=c_idx, value=value)
            cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                               top=Side(style='thin'), bottom=Side(style='thin'))
            
            # Apply formatting
            if c_idx in [4, 7]:  # Sales columns
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='right')
            elif c_idx in [5, 8]:  # QTY columns
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal='right')
            elif c_idx in [6, 9]:  # GP% columns
                cell.number_format = '0.00%'
                cell.alignment = Alignment(horizontal='right')
            elif c_idx == 1:  # Rank
                cell.alignment = Alignment(horizontal='center')
            elif c_idx == 2:  # Code
                cell.alignment = Alignment(horizontal='left')
            else:  # Item name
                cell.alignment = Alignment(horizontal='left')
    
    # Set column widths
    widths = [5, 12, 50, 15, 12, 12, 15, 12, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=2, column=i).column_letter].width = w
    
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 35
    
    # Create charts
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 8))
    fig.suptitle(f'Top 20 SKUs - {report_type} Performance ({month_name} {target_year})', 
                 fontsize=14, fontweight='bold')
    
    # Extract data for charts (top 20)
    skus_list = [f"{row['code'][:10]}..." if len(row['code']) > 10 else row['code'] 
                 for _, row in sku_df.head(20).iterrows()]
    sales_current = sku_df[f'{target_year}_{metric_lower}_sales'].head(20).tolist()
    sales_previous = sku_df[f'{comparison_year}_{metric_lower}_sales'].head(20).tolist()
    gp_current = sku_df[f'{target_year}_{metric_lower}_gp'].head(20).tolist()
    gp_previous = sku_df[f'{comparison_year}_{metric_lower}_gp'].head(20).tolist()
    
    # Chart 1: Sales Comparison
    y_pos = np.arange(len(skus_list))
    width = 0.35
    
    bars1 = ax1.barh(y_pos - width/2, sales_current, width, label=str(target_year), 
                    color='#2E86AB', alpha=0.8, edgecolor='black', linewidth=0.8)
    bars2 = ax1.barh(y_pos + width/2, sales_previous, width, label=str(comparison_year), 
                    color='#A23B72', alpha=0.8, edgecolor='black', linewidth=0.8)
    
    ax1.set_xlabel('Sales Amount ($)', fontsize=10, fontweight='bold')
    ax1.set_ylabel('SKU Code', fontsize=10, fontweight='bold')
    ax1.set_title(f'{report_type} Gross Sales', fontsize=11, fontweight='bold')
    ax1.set_yticks(y_pos)
    ax1.set_yticklabels(skus_list, fontsize=7)
    ax1.legend(fontsize=9)
    ax1.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x/1000:.0f}K'))
    ax1.grid(axis='x', alpha=0.3, linestyle='--')
    ax1.invert_yaxis()
    
    # Chart 2: GP% Comparison
    bars3 = ax2.barh(y_pos - width/2, gp_current, width, label=str(target_year), 
                    color='#2E86AB', alpha=0.8, edgecolor='black', linewidth=0.8)
    bars4 = ax2.barh(y_pos + width/2, gp_previous, width, label=str(comparison_year), 
                    color='#A23B72', alpha=0.8, edgecolor='black', linewidth=0.8)
    
    ax2.set_xlabel('GP %', fontsize=10, fontweight='bold')
    ax2.set_title(f'{report_type} Gross Profit %', fontsize=11, fontweight='bold')
    ax2.set_yticks(y_pos)
    ax2.set_yticklabels(skus_list, fontsize=7)
    ax2.legend(fontsize=9)
    ax2.grid(axis='x', alpha=0.3, linestyle='--')
    ax2.invert_yaxis()
    
    plt.tight_layout()
    
    # Save chart to bytes
    img_buffer = io.BytesIO()
    fig.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
    img_buffer.seek(0)
    plt.close(fig)
    
    # Add chart to new sheet
    ws_chart = wb.create_sheet(f"{report_type} Charts")
    img = XLImage(img_buffer)
    ws_chart.add_image(img, 'A1')
    
    # Save to bytes
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, sku_df

def get_sku_code_and_name(item_df):
    """Extract the original SKU code and a readable item name from the dataframe."""
    if item_df.empty:
        return '', ''

    first_row = item_df.iloc[0]

    # Extract code from Item Number column
    code = ''
    for col in item_df.columns:
        col_lower = str(col).lower()
        if 'item' in col_lower and any(k in col_lower for k in ['no', 'number', '#', 'code']):
            val = first_row[col]
            if pd.notna(val):
                val_str = str(val).strip()
                if val_str:
                    code = val_str.split()[0]
                    break

    if not code:
        raw_value = str(first_row.iloc[0]).strip()
        match = re.search(r'([A-Za-z]+\d+)', raw_value)
        if match:
            code = match.group(1)
        elif raw_value:
            code = raw_value.split()[0]

    # Extract name - look through all columns for a descriptive text
    name = ''
    # First try specific description columns
    for col in item_df.columns:
        col_lower = str(col).lower()
        if any(key in col_lower for key in ['description', 'item name', 'item desc']):
            val = first_row[col]
            if pd.notna(val):
                val_str = str(val).strip()
                if val_str and (not code or not val_str.upper().startswith(code.upper())):
                    name = val_str
                    break
    
    # If no name found, search all columns for a string that looks like a description
    # (contains spaces and is longer than the code)
    if not name:
        for col in item_df.columns:
            if col == 'Item Number':
                continue
            val = first_row[col]
            if pd.notna(val):
                val_str = str(val).strip()
                # Check if this looks like a product description
                if val_str and len(val_str) > 10 and ' ' in val_str and not val_str.replace('.','').replace(',','').isdigit():
                    # Make sure it's not just the code repeated
                    if not code or not val_str.upper().startswith(code.upper()):
                        name = val_str
                        break
                    # If it starts with code, extract the remainder
                    elif code and val_str.upper().startswith(code.upper()):
                        remainder = val_str[len(code):].strip(' -:_')
                        if remainder and len(remainder) > 3:
                            name = remainder
                            break

    if not name:
        raw_value = str(first_row.iloc[0]).strip()
        if code and raw_value.upper().startswith(code.upper()):
            remainder = raw_value[len(code):].strip(' -:_')
            if remainder:
                name = remainder
        elif raw_value and raw_value != code:
            name = raw_value

    if not code:
        code = name or ''
    if not name:
        name = code

    return code.strip(), name.strip()

def create_brand_report(all_items, target_month, target_year, comparison_year, month_name, report_type='MTD'):
    """Create Top-10 Brand MTD or YTD Performance report (excludes MX brand)"""
    
    # Group items by brand code (exclude MX brand)
    brands = defaultdict(list)
    
    for item_key, item_df in all_items.items():
        brand_code = extract_brand_code(item_key)
        # Skip MX brand from reports
        if brand_code == 'MX':
            continue
        brands[brand_code].append((item_key, item_df))
    
    # Aggregate per brand
    brand_metrics = []
    for brand_code, items in brands.items():
        agg = {
            'brand_code': brand_code,
            'brand_name': None,
            f'{target_year}_mtd_sales': 0.0,
            f'{target_year}_mtd_cost': 0.0,
            f'{target_year}_ytd_sales': 0.0,
            f'{target_year}_ytd_cost': 0.0,
            f'{comparison_year}_mtd_sales': 0.0,
            f'{comparison_year}_mtd_cost': 0.0,
            f'{comparison_year}_ytd_sales': 0.0,
            f'{comparison_year}_ytd_cost': 0.0,
        }
        desc_sets = []
        
        for item_key, item_df in items:
            metrics_current = compute_item_year_metrics(item_df, target_year, target_month)
            metrics_previous = compute_item_year_metrics(item_df, comparison_year, target_month)
            
            agg[f'{target_year}_mtd_sales'] += metrics_current['mtd_sales']
            agg[f'{target_year}_mtd_cost'] += metrics_current['mtd_cost']
            agg[f'{target_year}_ytd_sales'] += metrics_current['ytd_sales']
            agg[f'{target_year}_ytd_cost'] += metrics_current['ytd_cost']
            agg[f'{comparison_year}_mtd_sales'] += metrics_previous['mtd_sales']
            agg[f'{comparison_year}_mtd_cost'] += metrics_previous['mtd_cost']
            agg[f'{comparison_year}_ytd_sales'] += metrics_previous['ytd_sales']
            agg[f'{comparison_year}_ytd_cost'] += metrics_previous['ytd_cost']
            
            desc = metrics_current.get('desc') or metrics_previous.get('desc')
            if desc:
                words = [w.strip().upper() for w in re.split(r'\s+', str(desc)) if w.strip()]
                first3 = words[:3]
                if first3:
                    desc_sets.append(set(first3))
        
        # Determine brand name from common words
        if desc_sets:
            common = set.intersection(*desc_sets) if len(desc_sets) > 1 else desc_sets[0]
            if common:
                sample_desc = None
                for _, df in items:
                    for col in df.columns:
                        if any(k in str(col).lower() for k in ['description','item','name']):
                            non_nulls = df[col].dropna().astype(str)
                            if len(non_nulls):
                                sample_desc = non_nulls.iloc[0]
                                break
                    if sample_desc:
                        break
                if sample_desc:
                    sample_words = [w.strip() for w in re.split(r'\s+', sample_desc) if w.strip()]
                    brand_words = [w for w in sample_words[:3] if w.upper() in common]
                    brand_name = ' '.join(brand_words)
                else:
                    brand_name = ' '.join(list(common))
            else:
                brand_name = brand_code
        else:
            brand_name = brand_code
        
        agg['brand_name'] = brand_name
        
        # Compute GP%
        def compute_gp(sales, cost):
            return ((sales - cost) / sales * 100) if sales > 0 else 0.0
        
        agg[f'{target_year}_mtd_gp'] = compute_gp(agg[f'{target_year}_mtd_sales'], agg[f'{target_year}_mtd_cost'])
        agg[f'{target_year}_ytd_gp'] = compute_gp(agg[f'{target_year}_ytd_sales'], agg[f'{target_year}_ytd_cost'])
        agg[f'{comparison_year}_mtd_gp'] = compute_gp(agg[f'{comparison_year}_mtd_sales'], agg[f'{comparison_year}_mtd_cost'])
        agg[f'{comparison_year}_ytd_gp'] = compute_gp(agg[f'{comparison_year}_ytd_sales'], agg[f'{comparison_year}_ytd_cost'])
        
        # % Achieved
        agg['mtd_achieved_pct'] = (agg[f'{target_year}_mtd_sales'] / agg[f'{comparison_year}_mtd_sales'] * 100) if agg[f'{comparison_year}_mtd_sales'] > 0 else 0.0
        agg['ytd_achieved_pct'] = (agg[f'{target_year}_ytd_sales'] / agg[f'{comparison_year}_ytd_sales'] * 100) if agg[f'{comparison_year}_ytd_sales'] > 0 else 0.0
        
        brand_metrics.append(agg)
    
    # Convert to DataFrame and get top 10
    bm_df = pd.DataFrame(brand_metrics)
    if bm_df.empty:
        return None
    
    # Sort by MTD or YTD based on report type
    sort_key = f'{target_year}_{report_type.lower()}_sales'
    bm_df = bm_df.sort_values(by=sort_key, ascending=False).head(10)
    
    # Create Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.title = f'Top 10 {report_type} Brands'
    
    # Title
    ws.merge_cells('A1:I1')
    tcell = ws['A1']
    tcell.value = f'Top 10 {report_type} Brand Performance - {month_name} {target_year}'
    tcell.font = Font(size=14, bold=True)
    tcell.alignment = Alignment(horizontal='center')
    
    # Header row styling
    header_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    green_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    blue_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    
    headers = ['', 'Brand', 
               f'{target_year} {report_type}\nGross Sales', f'{target_year} {report_type}\nGP%',
               f'{comparison_year} {report_type}\nGross Sales', f'{comparison_year} {report_type}\nGP%',
               f'{target_year} {report_type}\nBudget ($)', 
               f'% Achieved vs\n{comparison_year}', '% Achieved vs\nBudget']
    
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=c_idx, value=h)
        cell.font = Font(bold=True, size=10)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                           top=Side(style='thin'), bottom=Side(style='thin'))
        
        # Color code headers
        if c_idx in [3, 4]:  # Current year columns
            cell.fill = header_fill
        elif c_idx in [5, 6]:  # Previous year columns
            cell.fill = green_fill
        elif c_idx in [7, 8, 9]:  # Budget/achieved columns
            cell.fill = blue_fill
    
    # Write data rows
    metric_lower = report_type.lower()
    for r_idx, row in bm_df.reset_index(drop=True).iterrows():
        rank = r_idx + 1
        brand_name = row['brand_name']
        curr_sales = row[f'{target_year}_{metric_lower}_sales']
        curr_gp = row[f'{target_year}_{metric_lower}_gp'] / 100
        prev_sales = row[f'{comparison_year}_{metric_lower}_sales']
        prev_gp = row[f'{comparison_year}_{metric_lower}_gp'] / 100
        achieved = row[f'{metric_lower}_achieved_pct'] / 100
        
        row_data = [rank, brand_name, curr_sales, curr_gp, prev_sales, prev_gp, None, achieved, None]
        
        r_idx = r_idx + 3  # Adjust for Excel row numbering (header is row 2, data starts at row 3)
        
        for c_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                               top=Side(style='thin'), bottom=Side(style='thin'))
            
            # Apply formatting
            if c_idx in [3, 5, 7]:  # Sales columns
                cell.number_format = '$#,##0.00'
                cell.alignment = Alignment(horizontal='right')
            elif c_idx in [4, 6, 8, 9]:  # Percentage columns
                cell.number_format = '0.00%'
                cell.alignment = Alignment(horizontal='right')
            elif c_idx == 1:  # Rank
                cell.alignment = Alignment(horizontal='center')
            else:  # Brand name
                cell.alignment = Alignment(horizontal='left')
    
    # Set column widths
    widths = [6, 22, 18, 12, 18, 12, 18, 14, 16]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=2, column=i).column_letter].width = w
    
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 35
    
    # Create charts
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
    fig.suptitle(f'Top 10 Brands - {report_type} Performance ({month_name} {target_year})', 
                 fontsize=14, fontweight='bold')
    
    # Extract data for charts
    brands_list = bm_df['brand_name'].head(10).tolist()
    metric_lower = report_type.lower()
    sales_current = bm_df[f'{target_year}_{metric_lower}_sales'].head(10).tolist()
    sales_previous = bm_df[f'{comparison_year}_{metric_lower}_sales'].head(10).tolist()
    gp_current = bm_df[f'{target_year}_{metric_lower}_gp'].head(10).tolist()
    gp_previous = bm_df[f'{comparison_year}_{metric_lower}_gp'].head(10).tolist()
    
    # Chart 1: Sales Comparison
    y_pos = np.arange(len(brands_list))
    width = 0.35
    
    bars1 = ax1.barh(y_pos - width/2, sales_current, width, label=str(target_year), 
                    color='#2E86AB', alpha=0.8, edgecolor='black', linewidth=1)
    bars2 = ax1.barh(y_pos + width/2, sales_previous, width, label=str(comparison_year), 
                    color='#A23B72', alpha=0.8, edgecolor='black', linewidth=1)
    
    ax1.set_xlabel('Sales Amount ($)', fontsize=10, fontweight='bold')
    ax1.set_ylabel('Brand', fontsize=10, fontweight='bold')
    ax1.set_title(f'{report_type} Gross Sales', fontsize=11, fontweight='bold')
    ax1.set_yticks(y_pos)
    ax1.set_yticklabels(brands_list, fontsize=8)
    ax1.legend(fontsize=9)
    ax1.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x/1000:.0f}K'))
    ax1.grid(axis='x', alpha=0.3, linestyle='--')
    ax1.invert_yaxis()
    
    # Add value labels
    for bars in [bars1, bars2]:
        for bar in bars:
            width_val = bar.get_width()
            if width_val > 0:
                ax1.text(width_val, bar.get_y() + bar.get_height()/2.,
                        f'${width_val/1000:.0f}K',
                        ha='left', va='center', fontsize=7, fontweight='bold')
    
    # Chart 2: GP% Comparison
    bars3 = ax2.barh(y_pos - width/2, gp_current, width, label=str(target_year), 
                    color='#2E86AB', alpha=0.8, edgecolor='black', linewidth=1)
    bars4 = ax2.barh(y_pos + width/2, gp_previous, width, label=str(comparison_year), 
                    color='#A23B72', alpha=0.8, edgecolor='black', linewidth=1)
    
    ax2.set_xlabel('GP %', fontsize=10, fontweight='bold')
    ax2.set_title(f'{report_type} Gross Profit %', fontsize=11, fontweight='bold')
    ax2.set_yticks(y_pos)
    ax2.set_yticklabels(brands_list, fontsize=8)
    ax2.legend(fontsize=9)
    ax2.grid(axis='x', alpha=0.3, linestyle='--')
    ax2.invert_yaxis()
    
    # Add value labels
    for bars in [bars3, bars4]:
        for bar in bars:
            width_val = bar.get_width()
            if width_val > 0:
                ax2.text(width_val, bar.get_y() + bar.get_height()/2.,
                        f'{width_val:.1f}%',
                        ha='left', va='center', fontsize=7, fontweight='bold')
    
    plt.tight_layout()
    
    # Save chart to bytes
    img_buffer = io.BytesIO()
    fig.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
    img_buffer.seek(0)
    plt.close(fig)
    
    # Add chart to new sheet
    ws_chart = wb.create_sheet(f"{report_type} Charts")
    img = XLImage(img_buffer)
    ws_chart.add_image(img, 'A1')
    
    # Save to bytes
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, bm_df, len(brands)

# Main processing
if uploaded_file is not None:
    try:
        with st.spinner("Processing your file..."):
            # Read Excel file
            excel_file = pd.ExcelFile(uploaded_file)
            
            st.success(f"✅ File uploaded successfully! Found {len(excel_file.sheet_names)} sheet(s)")
            
            # Show available sheets
            with st.expander("📄 Available Sheets"):
                st.write(excel_file.sheet_names)
            
            # Read and process all sheets
            dataframes = {}
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name, header=9)
                df = df.dropna(axis=1, how='all')

                unnamed_cols = [col for col in df.columns if isinstance(col, str) and col.startswith('Unnamed')]
                cols_to_drop = [col for col in unnamed_cols if df[col].isna().all()]
                if cols_to_drop:
                    df = df.drop(columns=cols_to_drop)

                if 'Unnamed: 0' in df.columns and 'Item Number' not in df.columns:
                    df = df.rename(columns={'Unnamed: 0': 'Item Number'})
                dataframes[sheet_name] = df
            
            # Split by items
            all_items = {}
            for sheet_name, df in dataframes.items():
                items = split_by_items(df)
                for item_num, item_df in items.items():
                    key = f"{sheet_name}_{item_num}"
                    all_items[key] = item_df
            
            st.info(f"📦 Extracted {len(all_items)} items from the file")
            
            # Calculate metrics
            results_current = calculate_sales_metrics(all_items, target_month, target_year)
            results_previous = calculate_sales_metrics(all_items, target_month, comparison_year)
            
            # Display results
            st.header(f"📊 Sales Summary - {month_name} {target_year}")
            
            # Create columns for metrics
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric(
                    "MTD Gross Sales", 
                    f"${results_current['MTD Gross Sales']:,.2f}",
                    f"{((results_current['MTD Gross Sales'] / results_previous['MTD Gross Sales'] - 1) * 100):.1f}%" if results_previous['MTD Gross Sales'] > 0 else "N/A"
                )
            
            with col2:
                st.metric(
                    "MTD GP%", 
                    f"{results_current['MTD GP%']:.2f}%",
                    f"{(results_current['MTD GP%'] - results_previous['MTD GP%']):.2f}%"
                )
            
            with col3:
                st.metric(
                    "YTD Gross Sales", 
                    f"${results_current['YTD Gross Sales']:,.2f}",
                    f"{((results_current['YTD Gross Sales'] / results_previous['YTD Gross Sales'] - 1) * 100):.1f}%" if results_previous['YTD Gross Sales'] > 0 else "N/A"
                )
            
            with col4:
                st.metric(
                    "YTD GP%", 
                    f"{results_current['YTD GP%']:.2f}%",
                    f"{(results_current['YTD GP%'] - results_previous['YTD GP%']):.2f}%"
                )
            
            # Create summary table
            st.subheader("📋 Detailed Summary Table")
            
            mtd_achieved = (results_current['MTD Gross Sales'] / results_previous['MTD Gross Sales'] * 100) if results_previous['MTD Gross Sales'] > 0 else 0
            ytd_achieved = (results_current['YTD Gross Sales'] / results_previous['YTD Gross Sales'] * 100) if results_previous['YTD Gross Sales'] > 0 else 0
            mtd_gp_achieved = "0%" if results_previous['MTD GP%'] == 0 else f"{(results_current['MTD GP%'] / results_previous['MTD GP%'] * 100):.0f}%"
            ytd_gp_achieved = "0%" if results_previous['YTD GP%'] == 0 else f"{(results_current['YTD GP%'] / results_previous['YTD GP%'] * 100):.0f}%"
            
            summary_data = {
                'Date': [str(target_year), str(comparison_year), '%Achieved', f'{target_year} Budget', '% Achieved'],
                'MTD Gross Sales': [
                    f"$ {results_current['MTD Gross Sales']:,.2f}",
                    f"$ {results_previous['MTD Gross Sales']:,.2f}",
                    f"{mtd_achieved:.0f}%",
                    "",
                    "0%"
                ],
                'MTD GP%': [
                    f"{results_current['MTD GP%']:.2f}%",
                    f"{results_previous['MTD GP%']:.2f}%",
                    mtd_gp_achieved,
                    "",
                    "0%"
                ],
                'YTD Gross Sales': [
                    f"$ {results_current['YTD Gross Sales']:,.2f}",
                    f"$ {results_previous['YTD Gross Sales']:,.2f}",
                    f"{ytd_achieved:.0f}%",
                    "",
                    "0%"
                ],
                'YTD GP%': [
                    f"{results_current['YTD GP%']:.2f}%",
                    f"{results_previous['YTD GP%']:.2f}%",
                    ytd_gp_achieved,
                    "",
                    "0%"
                ]
            }
            
            summary_df = pd.DataFrame(summary_data)
            st.dataframe(summary_df, use_container_width=True)
            
            # Generate Excel reports
            st.subheader("📥 Download Reports")
            
            # First row: Summary and Brand reports
            st.markdown("### 📊 Summary & Brand Reports")
            col_download1, col_download2, col_download3 = st.columns(3)
            
            with col_download1:
                st.markdown("#### Sales Summary Report")
                excel_output = create_excel_report(
                    results_current, 
                    results_previous, 
                    target_month, 
                    target_year, 
                    comparison_year,
                    month_name
                )
                
                st.download_button(
                    label="📥 Sales Summary",
                    data=excel_output,
                    file_name=f"Sales_Summary_Report_{month_name}_{target_year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                st.caption("Summary table & charts")
            
            with col_download2:
                st.markdown("#### Top 10 MTD Brands")
                brand_mtd_output, brand_mtd_df, total_brands = create_brand_report(
                    all_items,
                    target_month,
                    target_year,
                    comparison_year,
                    month_name,
                    'MTD'
                )
                
                if brand_mtd_output:
                    st.download_button(
                        label="📥 Top 10 MTD Brands",
                        data=brand_mtd_output,
                        file_name=f"Top10_Brand_MTD_Report_{month_name}_{target_year}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    st.caption(f"Top {len(brand_mtd_df)} by MTD sales")
                else:
                    st.warning("Could not generate MTD brand report")
                    total_brands = 0
            
            with col_download3:
                st.markdown("#### Top 10 YTD Brands")
                brand_ytd_output, brand_ytd_df, _ = create_brand_report(
                    all_items,
                    target_month,
                    target_year,
                    comparison_year,
                    month_name,
                    'YTD'
                )
                
                if brand_ytd_output:
                    st.download_button(
                        label="📥 Top 10 YTD Brands",
                        data=brand_ytd_output,
                        file_name=f"Top10_Brand_YTD_Report_{month_name}_{target_year}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    st.caption(f"Top {len(brand_ytd_df)} by YTD sales")
                else:
                    st.warning("Could not generate YTD brand report")
            
            # Second row: SKU reports
            st.markdown("### 🏷️ SKU Performance Reports")
            col_download4, col_download5, col_empty = st.columns(3)
            
            with col_download4:
                st.markdown("#### Top 20 MTD SKUs")
                sku_mtd_output, sku_mtd_df = create_sku_report(
                    all_items,
                    target_month,
                    target_year,
                    comparison_year,
                    month_name,
                    'MTD'
                )
                
                if sku_mtd_output:
                    st.download_button(
                        label="📥 Top 20 MTD SKUs",
                        data=sku_mtd_output,
                        file_name=f"Top20_SKU_MTD_Report_{month_name}_{target_year}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    st.caption(f"Top {len(sku_mtd_df)} SKUs by MTD sales")
                else:
                    st.warning("Could not generate MTD SKU report")
            
            with col_download5:
                st.markdown("#### Top 20 YTD SKUs")
                sku_ytd_output, sku_ytd_df = create_sku_report(
                    all_items,
                    target_month,
                    target_year,
                    comparison_year,
                    month_name,
                    'YTD'
                )
                
                if sku_ytd_output:
                    st.download_button(
                        label="📥 Top 20 YTD SKUs",
                        data=sku_ytd_output,
                        file_name=f"Top20_SKU_YTD_Report_{month_name}_{target_year}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    st.caption(f"Top {len(sku_ytd_df)} SKUs by YTD sales")
                else:
                    st.warning("Could not generate YTD SKU report")
            
            st.success("✅ All reports generated successfully! Click the buttons above to download.")
            
            # Show brand summary previews
            if brand_mtd_output:
                st.subheader("🏆 Top 10 Brands Preview")
                
                tab1, tab2 = st.tabs(["MTD Performance", "YTD Performance"])
                
                with tab1:
                    preview_mtd = []
                    for idx, row in brand_mtd_df.reset_index(drop=True).iterrows():
                        preview_mtd.append({
                            'Rank': idx + 1,
                            'Brand': row['brand_name'],
                            f'{target_year} MTD Sales': f"${row[f'{target_year}_mtd_sales']:,.2f}",
                            f'{target_year} GP%': f"{row[f'{target_year}_mtd_gp']:.2f}%",
                            f'{comparison_year} MTD Sales': f"${row[f'{comparison_year}_mtd_sales']:,.2f}",
                            f'{comparison_year} GP%': f"{row[f'{comparison_year}_mtd_gp']:.2f}%",
                            '% Achieved': f"{row['mtd_achieved_pct']:.2f}%"
                        })
                    
                    preview_mtd_df = pd.DataFrame(preview_mtd)
                    st.dataframe(preview_mtd_df, use_container_width=True, hide_index=True)
                
                with tab2:
                    if brand_ytd_output:
                        preview_ytd = []
                        for idx, row in brand_ytd_df.reset_index(drop=True).iterrows():
                            preview_ytd.append({
                                'Rank': idx + 1,
                                'Brand': row['brand_name'],
                                f'{target_year} YTD Sales': f"${row[f'{target_year}_ytd_sales']:,.2f}",
                                f'{target_year} GP%': f"{row[f'{target_year}_ytd_gp']:.2f}%",
                                f'{comparison_year} YTD Sales': f"${row[f'{comparison_year}_ytd_sales']:,.2f}",
                                f'{comparison_year} GP%': f"{row[f'{comparison_year}_ytd_gp']:.2f}%",
                                '% Achieved': f"{row['ytd_achieved_pct']:.2f}%"
                            })
                        
                        preview_ytd_df = pd.DataFrame(preview_ytd)
                        st.dataframe(preview_ytd_df, use_container_width=True, hide_index=True)
            
            # Show SKU summary previews
            if sku_mtd_output:
                st.subheader("🏷️ Top 20 SKUs Preview")
                
                tab3, tab4 = st.tabs(["MTD SKU Performance", "YTD SKU Performance"])
                
                with tab3:
                    preview_sku_mtd = []
                    for idx, row in sku_mtd_df.reset_index(drop=True).iterrows():
                        preview_sku_mtd.append({
                            'Rank': idx + 1,
                            'Code': row['code'],
                            'Item Name': row['name'][:50] + '...' if len(row['name']) > 50 else row['name'],
                            f'{target_year} MTD Sales': f"${row[f'{target_year}_mtd_sales']:,.2f}",
                            f'{target_year} GP%': f"{row[f'{target_year}_mtd_gp']:.2f}%",
                            f'{comparison_year} MTD Sales': f"${row[f'{comparison_year}_mtd_sales']:,.2f}",
                            f'{comparison_year} GP%': f"{row[f'{comparison_year}_mtd_gp']:.2f}%"
                        })
                    
                    preview_sku_mtd_df = pd.DataFrame(preview_sku_mtd)
                    st.dataframe(preview_sku_mtd_df, use_container_width=True, hide_index=True)
                
                with tab4:
                    if sku_ytd_output:
                        preview_sku_ytd = []
                        for idx, row in sku_ytd_df.reset_index(drop=True).iterrows():
                            preview_sku_ytd.append({
                                'Rank': idx + 1,
                                'Code': row['code'],
                                'Item Name': row['name'][:50] + '...' if len(row['name']) > 50 else row['name'],
                                f'{target_year} YTD Sales': f"${row[f'{target_year}_ytd_sales']:,.2f}",
                                f'{target_year} GP%': f"{row[f'{target_year}_ytd_gp']:.2f}%",
                                f'{comparison_year} YTD Sales': f"${row[f'{comparison_year}_ytd_sales']:,.2f}",
                                f'{comparison_year} GP%': f"{row[f'{comparison_year}_ytd_gp']:.2f}%"
                            })
                        
                        preview_sku_ytd_df = pd.DataFrame(preview_sku_ytd)
                        st.dataframe(preview_sku_ytd_df, use_container_width=True, hide_index=True)
            
            # Additional info
            st.info(f"ℹ️ Processed {results_current['items_processed']} items with data for {target_year} | {total_brands} brands identified")
            
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.exception(e)
else:
    st.info("👆 Please upload an Excel file to get started")
    
    # Show sample information
    with st.expander("ℹ️ How to use this app"):
        st.markdown("""
        ### Instructions:
        1. **Upload your Excel file** (.xls or .xlsx format)
        2. **Configure settings** in the sidebar:
           - Select target month for MTD calculations
           - Select target year for current data
           - Select comparison year for year-over-year analysis
        3. **View the results** in the dashboard
        4. **Download** the generated Excel report with:
           - Formatted summary table
           - Dashboard chart (4-panel visualization)
           - Comparison chart (year-over-year)
        
        ### Expected File Format:
        - Header row should be at row 9 (0-indexed)
        - Should contain columns: Year, Period, Sales Amount, Cost of Sales
        - Items identified by alphanumeric codes
        - "Item Total:" rows separate different items
        """)

# Footer
st.markdown("---")
st.markdown("Built with Streamlit 🎈 | Sales Report Generator v1.0")
