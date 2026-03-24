"""
=============================================================
PROJECT 1: Sales Performance & Data Quality Analysis
Tools: Python (Pandas, NumPy, OpenPyXL, XlsxWriter)
Author: Vishal Khare
=============================================================
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import warnings
warnings.filterwarnings('ignore')

# ─────────────────────────────────────────
# STEP 1: LOAD RAW DATA
# ─────────────────────────────────────────
print("=" * 55)
print("  SALES PERFORMANCE & DATA QUALITY ANALYSIS")
print("=" * 55)

df_raw = pd.read_csv('sales_raw_data.csv')
print(f"\n[1] Raw dataset loaded: {df_raw.shape[0]} rows × {df_raw.shape[1]} columns")

# ─────────────────────────────────────────
# STEP 2: DATA QUALITY REPORT (Before Cleaning)
# ─────────────────────────────────────────
print("\n[2] DATA QUALITY REPORT (Before Cleaning)")
print("-" * 40)
print(f"  Total Records     : {len(df_raw)}")
print(f"  Duplicate Rows    : {df_raw.duplicated().sum()}")
print(f"  Missing Values    :")
for col, count in df_raw.isnull().sum().items():
    if count > 0:
        print(f"    └─ {col:<15}: {count} ({count/len(df_raw)*100:.1f}%)")

# Inconsistent casing check
casing_issues = df_raw['Region'].apply(lambda x: x != x.strip().title() if pd.notnull(x) else False).sum()
print(f"  Casing Issues     : {casing_issues} rows (Region column)")

# ─────────────────────────────────────────
# STEP 3: DATA CLEANING
# ─────────────────────────────────────────
print("\n[3] CLEANING DATA...")

df = df_raw.copy()

# Fix casing inconsistencies
df['Region'] = df['Region'].str.strip().str.title()
df['Category'] = df['Category'].str.strip().str.title()

# Fill missing Product with 'Unknown'
df['Product'] = df['Product'].fillna('Unknown')

# Fill missing Revenue using Quantity * Unit_Price * (1 - Discount)
mask = df['Revenue'].isnull()
df.loc[mask, 'Revenue'] = (
    df.loc[mask, 'Quantity'] *
    df.loc[mask, 'Unit_Price'] *
    (1 - df.loc[mask, 'Discount'])
).round(2)

# Parse dates
df['Order_Date'] = pd.to_datetime(df['Order_Date'])
df['Month'] = df['Order_Date'].dt.month_name()
df['Month_Num'] = df['Order_Date'].dt.month
df['Quarter'] = df['Order_Date'].dt.to_period('Q').astype(str)

# Remove duplicates
before = len(df)
df.drop_duplicates(inplace=True)
print(f"  ✔ Removed {before - len(df)} duplicate rows")
print(f"  ✔ Fixed {casing_issues} region casing issues")
print(f"  ✔ Filled {mask.sum()} missing Revenue values")
print(f"  ✔ Filled {df_raw['Product'].isnull().sum()} missing Product values")
print(f"  ✔ Clean dataset: {df.shape[0]} rows")

# ─────────────────────────────────────────
# STEP 4: ANALYSIS
# ─────────────────────────────────────────
print("\n[4] RUNNING ANALYSIS...")

# Regional Performance
regional = df.groupby('Region').agg(
    Total_Revenue=('Revenue', 'sum'),
    Total_Orders=('Order_ID', 'count'),
    Avg_Order_Value=('Revenue', 'mean'),
    Avg_Discount=('Discount', 'mean')
).round(2).sort_values('Total_Revenue', ascending=False)
regional['Revenue_Share_%'] = (regional['Total_Revenue'] / regional['Total_Revenue'].sum() * 100).round(1)

# Category Performance
category = df.groupby('Category').agg(
    Total_Revenue=('Revenue', 'sum'),
    Total_Orders=('Order_ID', 'count'),
    Avg_Order_Value=('Revenue', 'mean')
).round(2).sort_values('Total_Revenue', ascending=False)

# Monthly Trend
monthly = df.groupby(['Month_Num', 'Month']).agg(
    Total_Revenue=('Revenue', 'sum'),
    Total_Orders=('Order_ID', 'count')
).round(2).reset_index().sort_values('Month_Num')
monthly['MoM_Growth_%'] = monthly['Total_Revenue'].pct_change().mul(100).round(1)

# Top 10 Products
top_products = df.groupby('Product')['Revenue'].sum().sort_values(ascending=False).head(10).reset_index()
top_products.columns = ['Product', 'Total_Revenue']

# Customer Revenue
customer = df.groupby('Customer_ID')['Revenue'].sum().sort_values(ascending=False).reset_index()
customer.columns = ['Customer_ID', 'Total_Revenue']
customer['Segment'] = pd.cut(
    customer['Total_Revenue'],
    bins=[0, 20000, 60000, float('inf')],
    labels=['Regular', 'Mid Value', 'High Value']
)

print("  ✔ Regional analysis complete")
print("  ✔ Category analysis complete")
print("  ✔ Monthly trend analysis complete")
print("  ✔ Customer segmentation complete")

# ─────────────────────────────────────────
# STEP 5: KEY INSIGHTS SUMMARY
# ─────────────────────────────────────────
print("\n[5] KEY INSIGHTS")
print("-" * 40)
print(f"  Total Revenue     : ₹{df['Revenue'].sum():>12,.2f}")
print(f"  Total Orders      : {len(df):>12,}")
print(f"  Avg Order Value   : ₹{df['Revenue'].mean():>12,.2f}")
print(f"  Top Region        : {regional.index[0]} ({regional['Revenue_Share_%'].iloc[0]}% share)")
print(f"  Top Category      : {category.index[0]}")
print(f"  Best Month        : {monthly.loc[monthly['Total_Revenue'].idxmax(), 'Month']}")
high_value_count = (customer['Segment'] == 'High Value').sum()
print(f"  High Value Cust.  : {high_value_count} customers")

# ─────────────────────────────────────────
# STEP 6: VISUALIZATIONS
# ─────────────────────────────────────────
print("\n[6] GENERATING CHARTS...")

fig, axes = plt.subplots(2, 2, figsize=(14, 10))
fig.patch.set_facecolor('#0f1117')
for ax in axes.flat:
    ax.set_facecolor('#1a1d2e')

colors_main = ['#00d4ff', '#7c3aed', '#f59e0b', '#10b981']
accent = '#00d4ff'

# Chart 1: Regional Revenue
ax1 = axes[0, 0]
bars = ax1.bar(regional.index, regional['Total_Revenue'], color=colors_main, edgecolor='none', width=0.5)
ax1.set_title('Revenue by Region', color='white', fontsize=13, fontweight='bold', pad=12)
ax1.set_ylabel('Revenue (₹)', color='#aaaaaa', fontsize=9)
ax1.tick_params(colors='#aaaaaa', labelsize=9)
ax1.spines[:].set_visible(False)
ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'₹{x/1e6:.1f}M'))
for bar, val in zip(bars, regional['Revenue_Share_%']):
    ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1000,
             f'{val}%', ha='center', va='bottom', color='white', fontsize=8, fontweight='bold')

# Chart 2: Category Breakdown
ax2 = axes[0, 1]
wedges, texts, autotexts = ax2.pie(
    category['Total_Revenue'],
    labels=category.index,
    autopct='%1.1f%%',
    colors=['#00d4ff','#7c3aed','#f59e0b','#10b981','#ef4444'],
    startangle=90,
    wedgeprops={'edgecolor': '#0f1117', 'linewidth': 2}
)
for text in texts: text.set_color('#cccccc')
for at in autotexts: at.set_color('white'); at.set_fontweight('bold')
ax2.set_title('Revenue by Category', color='white', fontsize=13, fontweight='bold', pad=12)

# Chart 3: Monthly Trend
ax3 = axes[1, 0]
ax3.fill_between(monthly['Month'].str[:3], monthly['Total_Revenue'], alpha=0.3, color=accent)
ax3.plot(monthly['Month'].str[:3], monthly['Total_Revenue'], color=accent, linewidth=2.5, marker='o', markersize=5)
ax3.set_title('Monthly Revenue Trend', color='white', fontsize=13, fontweight='bold', pad=12)
ax3.set_ylabel('Revenue (₹)', color='#aaaaaa', fontsize=9)
ax3.tick_params(colors='#aaaaaa', labelsize=8)
ax3.spines[:].set_visible(False)
ax3.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'₹{x/1e6:.1f}M'))
plt.setp(ax3.xaxis.get_majorticklabels(), rotation=45, ha='right')

# Chart 4: Top 10 Products
ax4 = axes[1, 1]
top_products_sorted = top_products.sort_values('Total_Revenue')
colors_bar = ['#10b981' if r == top_products_sorted['Total_Revenue'].max() else '#00d4ff'
              for r in top_products_sorted['Total_Revenue']]
bars4 = ax4.barh(top_products_sorted['Product'], top_products_sorted['Total_Revenue'],
                  color=colors_bar, edgecolor='none', height=0.6)
ax4.set_title('Top 10 Products by Revenue', color='white', fontsize=13, fontweight='bold', pad=12)
ax4.tick_params(colors='#aaaaaa', labelsize=8)
ax4.spines[:].set_visible(False)
ax4.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'₹{x/1e3:.0f}K'))

plt.suptitle('Sales Performance Dashboard — 2024', color='white', fontsize=16,
             fontweight='bold', y=1.01)
plt.tight_layout(pad=2)
plt.savefig('/home/claude/sales_dashboard.png', dpi=150, bbox_inches='tight',
            facecolor='#0f1117')
plt.close()
print("  ✔ Dashboard chart saved: sales_dashboard.png")

# ─────────────────────────────────────────
# STEP 7: EXPORT TO EXCEL
# ─────────────────────────────────────────
print("\n[7] EXPORTING TO EXCEL...")

with pd.ExcelWriter('/home/claude/sales_analysis.xlsx', engine='xlsxwriter') as writer:
    wb = writer.book

    # Formats
    header_fmt = wb.add_format({'bold': True, 'bg_color': '#1e3a5f', 'font_color': 'white',
                                  'border': 1, 'align': 'center', 'valign': 'vcenter'})
    num_fmt = wb.add_format({'num_format': '₹#,##0.00', 'border': 1})
    pct_fmt = wb.add_format({'num_format': '0.0%', 'border': 1})
    text_fmt = wb.add_format({'border': 1})
    title_fmt = wb.add_format({'bold': True, 'font_size': 14, 'font_color': '#1e3a5f'})
    alt_fmt = wb.add_format({'bg_color': '#eaf2ff', 'border': 1})

    def write_sheet(df_write, sheet_name, title):
        df_write.to_excel(writer, sheet_name=sheet_name, startrow=2, index=True)
        ws = writer.sheets[sheet_name]
        ws.write(0, 0, title, title_fmt)
        ws.set_column('A:A', 20)
        ws.set_column('B:Z', 18)
        # Header row formatting
        for col_num, col in enumerate(df_write.columns):
            ws.write(2, col_num + 1, col, header_fmt)

    write_sheet(regional, 'Regional Analysis', '📍 Regional Performance Summary')
    write_sheet(category, 'Category Analysis', '📦 Category Performance Summary')
    monthly_out = monthly[['Month', 'Total_Revenue', 'Total_Orders', 'MoM_Growth_%']]
    monthly_out.to_excel(writer, sheet_name='Monthly Trends', startrow=2, index=False)
    ws_m = writer.sheets['Monthly Trends']
    ws_m.write(0, 0, '📅 Monthly Revenue Trends', title_fmt)
    top_products.to_excel(writer, sheet_name='Top Products', startrow=2, index=False)
    ws_p = writer.sheets['Top Products']
    ws_p.write(0, 0, '🏆 Top 10 Products by Revenue', title_fmt)
    df.to_excel(writer, sheet_name='Cleaned Data', index=False)
    df_raw.to_excel(writer, sheet_name='Raw Data', index=False)

print("  ✔ Excel file saved: sales_analysis.xlsx")
print("\n✅ PROJECT 1 COMPLETE!\n")
