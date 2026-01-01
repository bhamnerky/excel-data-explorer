#!/usr/bin/env python3
import pandas as pd
import duckdb
import warnings
warnings.filterwarnings('ignore')

print("🦆 Setting up DuckDB for Casing WIP Analysis")
print("="*80)

# Load Excel with proper headers
print("\n📥 Loading Excel file...")
file_path = 'FilesIn/casing.xlsx'
df = pd.read_excel(file_path, sheet_name='WIP - P10', header=1)
print(f"   ✓ Loaded {len(df):,} records with {len(df.columns)} columns")

# Clean up data types for DuckDB compatibility
print("\n🧹 Cleaning data types...")
date_cols = ['WIPMth', 'Start Month', 'MonthClosed',
             'Dispatcher Start Date', 'Dispatcher End Date']
for col in date_cols:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors='coerce')
        print(f"   ✓ Converted {col} to datetime")

# Create DuckDB connection (persistent to disk)
print("\n🔧 Creating DuckDB connection...")
conn = duckdb.connect('casing_analysis.duckdb')
print("   ✓ Connected to casing_analysis.duckdb")

# Register and create table
print("\n📊 Creating table in DuckDB...")
conn.execute("DROP TABLE IF EXISTS casing_wip")
conn.register('df_view', df)
conn.execute("CREATE TABLE casing_wip AS SELECT * FROM df_view")
print("   ✓ Table 'casing_wip' created successfully")

# Verify
row_count = conn.execute("SELECT COUNT(*) FROM casing_wip").fetchone()[0]
print(f"   ✓ Verified: {row_count:,} rows in table")

# Run test queries
print("\n🧪 Test Query 1: Contract Status Summary")
print("-"*80)
result = conn.execute("""
    SELECT "Contract Status",
           COUNT(*) as count,
           ROUND(SUM("Revised Contract")/1000000, 2) as contract_value_m,
           ROUND(SUM("Revenue To Date")/1000000, 2) as revenue_m,
           ROUND(AVG("Gross Profit %") * 100, 1) as avg_margin_pct
    FROM casing_wip
    WHERE "Contract Status" IS NOT NULL
    GROUP BY "Contract Status"
    ORDER BY contract_value_m DESC
""").df()
print(result.to_string(index=False))

print("\n🧪 Test Query 2: Top 10 Contracts by Revenue")
print("-"*80)
result = conn.execute("""
    SELECT Contract, 
           Description,
           "PM Name",
           ROUND("Revenue To Date"/1000000, 2) as revenue_m,
           ROUND("Gross Profit %" * 100, 1) as margin_pct,
           "Contract Status"
    FROM casing_wip
    WHERE "Revenue To Date" > 0
    ORDER BY "Revenue To Date" DESC
    LIMIT 10
""").df()
print(result.to_string(index=False))

print("\n🧪 Test Query 3: Regional Performance")
print("-"*80)
result = conn.execute("""
    SELECT Region,
           COUNT(*) as contracts,
           ROUND(SUM("Revenue To Date")/1000000, 2) as revenue_m,
           ROUND(AVG("Gross Profit %") * 100, 1) as avg_margin_pct
    FROM casing_wip
    WHERE Region IS NOT NULL AND Region != ''
    GROUP BY Region
    ORDER BY revenue_m DESC
""").df()
print(result.to_string(index=False))

conn.close()
print("\n" + "="*80)
print("✅ Database created and saved to: casing_analysis.duckdb")
print("\n📋 Quick Start:")
print("   import duckdb")
print("   conn = duckdb.connect('casing_analysis.duckdb')")
print("   result = conn.execute('SELECT * FROM casing_wip LIMIT 10').df()")
print("="*80)
