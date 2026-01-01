#!/usr/bin/env python3
"""
Pre-built Analytics Queries for WIP Analysis
Run business-focused queries on the DuckDB database
"""
import duckdb

conn = duckdb.connect('../wip_analysis.duckdb', read_only=True)


def run_query(query, title):
    print(f"\n{'='*80}")
    print(f"📊 {title}")
    print('='*80)
    result = conn.execute(query).df()
    print(result.to_string(index=False))
    print(f"\n✓ {len(result)} rows")
    return result


# Query 1: Overall Portfolio Health
run_query("""
    SELECT 
        "Contract Status",
        COUNT(*) as contracts,
        ROUND(SUM("Revenue To Date")/1000000, 2) as revenue_m,
        ROUND(SUM("Gross Profit")/1000000, 2) as profit_m,
        ROUND(AVG("Gross Profit %") * 100, 1) as avg_margin_pct,
        ROUND(AVG("% Complete") * 100, 1) as avg_complete_pct
    FROM wip
    WHERE "Contract Status" NOT IN ('InterCo Elim', 'ASC 606 Adjustment')
    GROUP BY "Contract Status"
    ORDER BY revenue_m DESC
""", "Portfolio Health by Contract Status")

# Query 2: Margin Distribution Analysis
run_query("""
    SELECT 
        CASE 
            WHEN "Gross Profit %" < 0 THEN 'Loss (< 0%)'
            WHEN "Gross Profit %" < 0.15 THEN 'Low (0-15%)'
            WHEN "Gross Profit %" < 0.30 THEN 'Medium (15-30%)'
            ELSE 'High (> 30%)'
        END as margin_bucket,
        COUNT(*) as contracts,
        ROUND(SUM("Revenue To Date")/1000000, 2) as revenue_m,
        ROUND(AVG("% Complete") * 100, 1) as avg_complete_pct
    FROM wip
    WHERE "Revenue To Date" > 0
    GROUP BY margin_bucket
    ORDER BY 
        CASE margin_bucket
            WHEN 'Loss (< 0%)' THEN 1
            WHEN 'Low (0-15%)' THEN 2
            WHEN 'Medium (15-30%)' THEN 3
            ELSE 4
        END
""", "Margin Distribution Analysis")

# Query 3: At-Risk Contracts
run_query("""
    SELECT 
        Contract,
        "Customer Name",
        Region,
        "PM Name",
        ROUND("Revenue To Date"/1000000, 2) as revenue_m,
        ROUND("Gross Profit %" * 100, 1) as margin_pct,
        ROUND("% Complete" * 100, 1) as complete_pct
    FROM wip
    WHERE "Contract Status" = 'Open'
      AND "Gross Profit %" < 0.15
      AND "Revenue To Date" > 100000
    ORDER BY "Revenue To Date" DESC
    LIMIT 20
""", "⚠️  At-Risk Contracts (Open, Low Margin, >$100K)")

# Query 4: Regional Performance Comparison
run_query("""
    SELECT 
        Region,
        COUNT(*) as contracts,
        ROUND(SUM("Revenue To Date")/1000000, 2) as revenue_m,
        ROUND(SUM("Gross Profit")/1000000, 2) as profit_m,
        ROUND(AVG("Gross Profit %") * 100, 1) as avg_margin_pct,
        ROUND(SUM("Backlog Revenue")/1000000, 2) as backlog_m
    FROM wip
    WHERE Region IS NOT NULL
      AND "Contract Status" NOT IN ('InterCo Elim')
    GROUP BY Region
    ORDER BY revenue_m DESC
""", "Regional Performance Comparison")

# Query 5: Top Customers by Revenue
run_query("""
    SELECT 
        "Customer Name",
        COUNT(*) as contracts,
        ROUND(SUM("Revenue To Date")/1000000, 2) as revenue_m,
        ROUND(AVG("Gross Profit %") * 100, 1) as avg_margin_pct
    FROM wip
    WHERE "Customer Name" IS NOT NULL
      AND "Revenue To Date" > 0
    GROUP BY "Customer Name"
    ORDER BY revenue_m DESC
    LIMIT 15
""", "Top 15 Customers by Revenue")

# Query 6: Service Type Analysis
run_query("""
    SELECT 
        ServiceType,
        COUNT(*) as contracts,
        ROUND(SUM("Revenue To Date")/1000000, 2) as revenue_m,
        ROUND(AVG("Gross Profit %") * 100, 1) as avg_margin_pct,
        ROUND(AVG("% Complete") * 100, 1) as avg_complete_pct
    FROM wip
    WHERE ServiceType IS NOT NULL
      AND "Revenue To Date" > 0
    GROUP BY ServiceType
    ORDER BY revenue_m DESC
""", "Service Type Performance")

# Query 7: PM Performance Leaderboard
run_query("""
    SELECT 
        "PM Name",
        COUNT(*) as contracts,
        ROUND(SUM("Revenue To Date")/1000000, 2) as revenue_m,
        ROUND(SUM("Gross Profit")/1000000, 2) as profit_m,
        ROUND(AVG("Gross Profit %") * 100, 1) as avg_margin_pct
    FROM wip
    WHERE "PM Name" IS NOT NULL
      AND "Revenue To Date" > 0
    GROUP BY "PM Name"
    HAVING COUNT(*) >= 5
    ORDER BY revenue_m DESC
    LIMIT 20
""", "Top 20 Project Managers (≥5 contracts)")

# Query 8: Large Projects (>$5M revenue)
run_query("""
    SELECT 
        Contract,
        Description,
        "Customer Name",
        Region,
        ROUND("Revenue To Date"/1000000, 2) as revenue_m,
        ROUND("Gross Profit %" * 100, 1) as margin_pct,
        ROUND("% Complete" * 100, 1) as complete_pct,
        "Contract Status"
    FROM wip
    WHERE "Revenue To Date" > 5000000
    ORDER BY "Revenue To Date" DESC
    LIMIT 20
""", "Large Projects (>$5M Revenue)")

conn.close()

print("\n" + "="*80)
print("✅ All queries completed successfully")
print("="*80)
print("\n💡 For custom queries, use: python3 custom_query.py")
