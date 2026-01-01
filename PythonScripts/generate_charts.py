#!/usr/bin/env python3
"""
Generate Charts for WIP Analysis Summary
Creates professional visualizations for markdown embedding
"""
import duckdb
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path

# Set professional style
sns.set_theme(style="whitegrid")
plt.rcParams['figure.figsize'] = (12, 6)
plt.rcParams['font.size'] = 10

# Create output directory
Path('../charts').mkdir(exist_ok=True)

conn = duckdb.connect('../wip_analysis.duckdb', read_only=True)

print("📊 Generating visualizations for WIP Analysis...")

# CHART 1: Regional Revenue Comparison
print("\n1. Regional revenue comparison...")
df = conn.execute("""
    SELECT Region, 
           SUM("Revenue To Date")/1000000 as revenue_m
    FROM wip 
    WHERE Region IS NOT NULL
      AND "Contract Status" NOT IN ('InterCo Elim')
    GROUP BY Region 
    ORDER BY revenue_m DESC
    LIMIT 10
""").df()

plt.figure(figsize=(12, 6))
plt.barh(df['Region'], df['revenue_m'], color='steelblue', edgecolor='navy')
plt.xlabel('Revenue ($M)', fontsize=12, fontweight='bold')
plt.title('Revenue by Region (Top 10)', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig('../charts/01_regional_revenue.png', dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Saved: ../charts/01_regional_revenue.png")

# CHART 2: Margin Distribution
print("2. Margin distribution...")
df = conn.execute("""
    SELECT "Gross Profit %" * 100 as margin_pct
    FROM wip 
    WHERE "Gross Profit %" IS NOT NULL
      AND "Revenue To Date" > 0
      AND "Gross Profit %" BETWEEN -0.5 AND 1.5
""").df()

plt.figure(figsize=(12, 6))
plt.hist(df['margin_pct'], bins=50, color='coral',
         edgecolor='black', alpha=0.7)
plt.xlabel('Margin (%)', fontsize=12, fontweight='bold')
plt.ylabel('Number of Contracts', fontsize=12, fontweight='bold')
plt.title('Contract Margin Distribution', fontsize=14, fontweight='bold')
median_val = df['margin_pct'].median()
plt.axvline(median_val, color='red', linestyle='--', linewidth=2,
            label=f'Median: {median_val:.1f}%')
plt.axvline(30, color='green', linestyle='--', linewidth=2, alpha=0.7,
            label='Target: 30%')
plt.axvline(0, color='darkred', linestyle='-', linewidth=2, alpha=0.7,
            label='Break-even')
plt.legend(fontsize=10)
plt.tight_layout()
plt.savefig('../charts/02_margin_distribution.png', dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Saved: ../charts/02_margin_distribution.png")

# CHART 3: Contract Status Composition
print("3. Contract status composition...")
df = conn.execute("""
    SELECT "Contract Status", COUNT(*) as count
    FROM wip 
    WHERE "Contract Status" NOT IN ('InterCo Elim', 'ASC 606 Adjustment')
    GROUP BY "Contract Status"
    ORDER BY count DESC
""").df()

plt.figure(figsize=(10, 8))
colors = plt.cm.Set3(range(len(df)))
explode = [0.05 if i == 0 else 0 for i in range(len(df))]
plt.pie(df['count'], labels=df['Contract Status'], autopct='%1.1f%%',
        startangle=90, colors=colors, explode=explode)
plt.title('Portfolio Composition by Contract Status',
          fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig('../charts/03_status_composition.png', dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Saved: ../charts/03_status_composition.png")

# CHART 4: Revenue vs Margin Scatter
print("4. Revenue vs margin correlation...")
df = conn.execute("""
    SELECT "Revenue To Date"/1000000 as revenue_m, 
           "Gross Profit %" * 100 as margin_pct,
           "Contract Status"
    FROM wip 
    WHERE "Revenue To Date" > 100000
      AND "Gross Profit %" BETWEEN -0.5 AND 1.0
      AND "Contract Status" IN ('Open', 'Soft-Closed')
    LIMIT 1000
""").df()

plt.figure(figsize=(12, 8))
colors_map = {'Open': 'blue', 'Soft-Closed': 'green'}
for status in df['Contract Status'].unique():
    subset = df[df['Contract Status'] == status]
    plt.scatter(subset['revenue_m'], subset['margin_pct'],
                alpha=0.5, label=status, s=30, c=colors_map.get(status, 'gray'))

plt.xlabel('Revenue ($M)', fontsize=12, fontweight='bold')
plt.ylabel('Gross Profit Margin (%)', fontsize=12, fontweight='bold')
plt.title('Revenue vs Margin Analysis', fontsize=14, fontweight='bold')
plt.axhline(y=30, color='green', linestyle='--',
            alpha=0.7, label='Target: 30%')
plt.axhline(y=0, color='red', linestyle='--', alpha=0.7, label='Break-even')
plt.legend(fontsize=10)
plt.grid(alpha=0.3)
plt.tight_layout()
plt.savefig('../charts/04_revenue_vs_margin.png', dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Saved: ../charts/04_revenue_vs_margin.png")

# CHART 5: Service Type Breakdown
print("5. Service type breakdown...")
df = conn.execute("""
    SELECT ServiceType,
           COUNT(*) as contracts,
           SUM("Revenue To Date")/1000000 as revenue_m
    FROM wip 
    WHERE ServiceType IS NOT NULL
      AND "Revenue To Date" > 0
      AND ServiceType NOT LIKE '%x000D%'
    GROUP BY ServiceType
    ORDER BY revenue_m DESC
    LIMIT 5
""").df()

fig, ax1 = plt.subplots(figsize=(12, 6))

x = range(len(df))
ax1.bar(x, df['contracts'], color='lightblue',
        edgecolor='navy', label='Contracts', alpha=0.7)
ax1.set_xlabel('Service Type', fontsize=12, fontweight='bold')
ax1.set_ylabel('Number of Contracts', fontsize=12,
               fontweight='bold', color='navy')
ax1.tick_params(axis='y', labelcolor='navy')
ax1.set_xticks(x)
ax1.set_xticklabels(df['ServiceType'], rotation=45, ha='right')

ax2 = ax1.twinx()
ax2.plot(x, df['revenue_m'], color='darkred', marker='o', linewidth=3,
         markersize=10, label='Revenue ($M)')
ax2.set_ylabel('Revenue ($M)', fontsize=12, fontweight='bold', color='darkred')
ax2.tick_params(axis='y', labelcolor='darkred')

plt.title('Service Type Performance', fontsize=14, fontweight='bold')
fig.tight_layout()
plt.savefig('../charts/05_service_type.png', dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Saved: ../charts/05_service_type.png")

# CHART 6: Completion Status for Open Contracts
print("6. Completion status for open contracts...")
df = conn.execute("""
    SELECT 
        CASE 
            WHEN "% Complete" < 0.25 THEN '0-25%'
            WHEN "% Complete" < 0.50 THEN '25-50%'
            WHEN "% Complete" < 0.75 THEN '50-75%'
            WHEN "% Complete" < 1.00 THEN '75-99%'
            ELSE '100%'
        END as completion_bucket,
        COUNT(*) as contracts,
        SUM("Revenue To Date")/1000000 as revenue_m
    FROM wip
    WHERE "Contract Status" = 'Open'
    GROUP BY completion_bucket
    ORDER BY 
        CASE completion_bucket
            WHEN '0-25%' THEN 1
            WHEN '25-50%' THEN 2
            WHEN '50-75%' THEN 3
            WHEN '75-99%' THEN 4
            ELSE 5
        END
""").df()

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))

# Contracts count
ax1.bar(df['completion_bucket'], df['contracts'],
        color='skyblue', edgecolor='navy')
ax1.set_xlabel('Completion Status', fontsize=11, fontweight='bold')
ax1.set_ylabel('Number of Contracts', fontsize=11, fontweight='bold')
ax1.set_title('Open Contracts by Completion', fontsize=12, fontweight='bold')
ax1.grid(axis='y', alpha=0.3)

# Revenue
ax2.bar(df['completion_bucket'], df['revenue_m'],
        color='orange', edgecolor='darkred')
ax2.set_xlabel('Completion Status', fontsize=11, fontweight='bold')
ax2.set_ylabel('Revenue ($M)', fontsize=11, fontweight='bold')
ax2.set_title('Revenue by Completion Level', fontsize=12, fontweight='bold')
ax2.grid(axis='y', alpha=0.3)

plt.tight_layout()
plt.savefig('../charts/06_completion_status.png', dpi=150, bbox_inches='tight')
plt.close()
print("   ✓ Saved: ../charts/06_completion_status.png")

conn.close()

print("\n✅ All charts generated successfully!")
print("   Location: ../charts/ directory")
print("   Ready for embedding in summary document")
