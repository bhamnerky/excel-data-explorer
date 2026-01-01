#!/usr/bin/env python3
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

print("🔍 Inspecting casing.xlsx structure...")
print("="*80)

file_path = 'FilesIn/casing.xlsx'

# First, get sheet names
xls = pd.ExcelFile(file_path)
print(f"\n📋 Available Sheets: {xls.sheet_names}")

# Inspect first sheet's raw structure to find header row
sheet_name = xls.sheet_names[0]
print(f"\n🔍 Inspecting sheet: '{sheet_name}'")
print("-"*80)

# Read first 10 rows without assuming header location
df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=10)

print("\n📊 First 10 rows (finding header row):")
for i in range(min(10, len(df_raw))):
    row_data = df_raw.iloc[i].tolist()[:8]  # Show first 8 columns
    # Truncate long values for readability
    row_display = [str(val)[:30] if pd.notna(
        val) else 'NaN' for val in row_data]
    print(f"Row {i}: {row_display}")

print("\n" + "="*80)
print("✓ Inspection complete - identify which row contains actual headers")
