#!/usr/bin/env python3
"""
DuckDB Setup for WIP (Work in Progress) Analysis
Loads casing.xlsx into a persistent DuckDB database for SQL querying
"""
import pandas as pd
import duckdb
from pathlib import Path

print("🦆 Setting up DuckDB for WIP Analysis")
print("="*80)

# Load Excel with proper headers
print("\n📥 Loading Excel file...")
file_path = '../FilesIn/casing.xlsx'
df = pd.read_excel(file_path, sheet_name='WIP - P10', header=1)
print(f"   ✓ Loaded {len(df):,} records with {len(df.columns)} columns")

# Clean up data types for DuckDB compatibility
print("\n🧹 Cleaning data types...")
date_cols = ['WIPMth', 'Start Month', 'MonthClosed', 'Start Date']
for col in date_cols:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors='coerce')

# Handle any problematic column names (spaces, special chars are OK in DuckDB)
print("   ✓ Date columns converted")

# Create DuckDB connection (persistent to disk)
print("\n🔧 Creating DuckDB connection...")
db_path = '../wip_analysis.duckdb'
conn = duckdb.connect(db_path)
print(f"   ✓ Connected to {db_path}")

# Register and create table
print("\n📊 Creating table in DuckDB...")
conn.execute("DROP TABLE IF EXISTS wip")
conn.register('df_view', df)
conn.execute("CREATE TABLE wip AS SELECT * FROM df_view")
print("   ✓ Table 'wip' created successfully")

# Verify
row_count = conn.execute("SELECT COUNT(*) FROM wip").fetchone()[0]
print(f"   ✓ Verified: {row_count:,} rows in table")

# Run a test query
print("\n🧪 Test Query: Regional Performance")
result = conn.execute("""
    SELECT Region, 
           COUNT(*) as contracts,
           ROUND(SUM("Revenue To Date")/1000000, 2) as revenue_m,
           ROUND(AVG("Gross Profit %") * 100, 1) as avg_margin_pct
    FROM wip
    WHERE Region IS NOT NULL
    GROUP BY Region
    ORDER BY revenue_m DESC
    LIMIT 10
""").df()
print(result.to_string(index=False))

conn.close()
print("\n✅ Database saved to: ../wip_analysis.duckdb")
print("\n💡 Next steps:")
print("   - Run pre-built queries: python3 PythonScripts/query_wip.py")
print("   - Custom queries: python3 PythonScripts/custom_query.py")
print("   - Interactive Python: import duckdb; conn = duckdb.connect('wip_analysis.duckdb')")
