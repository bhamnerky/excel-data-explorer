#!/usr/bin/env python3
"""
Comprehensive analysis of casing.xlsx dataset
"""
import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

file_path = 'FilesIn/casing.xlsx'
sheet = 'WIP - P10'

# Load with correct header row (Row 1)
df = pd.read_excel(file_path, sheet_name=sheet, header=1)

print('='*80)
print('📊 CASING DATASET OVERVIEW')
print('='*80)
print(f'Records: {len(df):,}')
print(f'Columns: {len(df.columns)}')
print(f'Sheet: {sheet}')

# Date ranges
if 'WIPMth' in df.columns:
    df['WIPMth'] = pd.to_datetime(df['WIPMth'], errors='coerce')
    valid_dates = df['WIPMth'].dropna()
    if len(valid_dates) > 0:
        print(
            f'Period: {valid_dates.min().strftime("%Y-%m-%d")} to {valid_dates.max().strftime("%Y-%m-%d")}')

# Contract Status Distribution
if 'Contract Status' in df.columns:
    print(f'\n📋 Contract Status Distribution:')
    status_dist = df['Contract Status'].value_counts()
    for status, count in status_dist.items():
        pct = (count / len(df)) * 100
        print(f'  {status:20s}: {count:5,} ({pct:5.1f}%)')

# Financial Summary
print(f'\n💰 Financial Summary:')
financial_cols = {
    'Revised Contract': 'Total Contract Value',
    'Total Billings': 'Total Billings',
    'Revenue To Date': 'Revenue To Date',
    'Costs To Date': 'Costs To Date',
    'Gross Profit': 'Gross Profit'
}

for col, label in financial_cols.items():
    if col in df.columns:
        total = df[col].sum()
        if abs(total) > 1_000_000:
            print(f'  {label:25s}: ${total/1_000_000:>10.2f}M')
        else:
            print(f'  {label:25s}: ${total:>13,.2f}')

# Key metrics
if 'Gross Profit %' in df.columns:
    avg_margin = df['Gross Profit %'].mean() * 100
    median_margin = df['Gross Profit %'].median() * 100
    print(f'\n📈 Margin Metrics:')
    print(f'  Average Margin: {avg_margin:6.1f}%')
    print(f'  Median Margin:  {median_margin:6.1f}%')

if '% Complete' in df.columns:
    avg_complete = df['% Complete'].mean() * 100
    print(f'  Average Completion: {avg_complete:6.1f}%')

# Top by Revenue
print(f'\n🏆 Top 10 Contracts by Revenue:')
if 'Revenue To Date' in df.columns and 'Description' in df.columns and 'Contract' in df.columns:
    top_revenue = df.nlargest(10, 'Revenue To Date')[
        ['Contract', 'Description', 'Revenue To Date']]
    if 'Gross Profit %' in df.columns:
        top_revenue = df.nlargest(10, 'Revenue To Date')[
            ['Contract', 'Description', 'Revenue To Date', 'Gross Profit %']]
        for idx, row in top_revenue.iterrows():
            contract = str(row['Contract'])
            desc = str(row['Description'])[:40]
            revenue = row['Revenue To Date']
            margin = row['Gross Profit %'] * \
                100 if pd.notna(row['Gross Profit %']) else 0
            print(
                f'  {contract:15s} {desc:42s} ${revenue/1_000_000:>7.2f}M ({margin:>5.1f}%)')
    else:
        for idx, row in top_revenue.iterrows():
            contract = str(row['Contract'])
            desc = str(row['Description'])[:40]
            revenue = row['Revenue To Date']
            print(f'  {contract:15s} {desc:42s} ${revenue/1_000_000:>7.2f}M')

# Regional breakdown
if 'Region' in df.columns and 'Revenue To Date' in df.columns:
    print(f'\n🌍 Top Regions by Revenue:')
    region_revenue = df.groupby('Region')['Revenue To Date'].sum(
    ).sort_values(ascending=False).head(10)
    total_revenue = df['Revenue To Date'].sum()
    for region, revenue in region_revenue.items():
        pct = (revenue / total_revenue) * 100 if total_revenue != 0 else 0
        print(f'  {str(region):25s}: ${revenue/1_000_000:>7.2f}M ({pct:>5.1f}%)')

# PM breakdown
if 'PM Name' in df.columns and 'Revenue To Date' in df.columns:
    print(f'\n👤 Top 10 Project Managers by Revenue:')
    pm_revenue = df.groupby('PM Name')['Revenue To Date'].sum(
    ).sort_values(ascending=False).head(10)
    for pm, revenue in pm_revenue.items():
        print(f'  {str(pm):30s}: ${revenue/1_000_000:>7.2f}M')

# Column overview
print(f'\n📑 Key Columns Available ({len(df.columns)} total):')
key_patterns = ['Contract', 'Revenue', 'Cost', 'Profit',
                'Margin', 'Complete', 'Status', 'Region', 'PM', 'Billing']
key_cols = [col for col in df.columns if any(
    pattern in col for pattern in key_patterns)]
for i, col in enumerate(key_cols[:25], 1):
    print(f'  {i:2d}. {col}')
if len(key_cols) > 25:
    print(f'  ... and {len(key_cols) - 25} more key columns')

print(f'\n✅ Analysis complete!')
