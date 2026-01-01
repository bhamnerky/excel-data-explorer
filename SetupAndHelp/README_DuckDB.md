# WIP Analysis - DuckDB Setup

## 📊 Dataset Overview

**File:** `FilesIn/casing.xlsx` (Sheet: WIP - P10)  
**Database:** `wip_analysis.duckdb`  
**Records:** 2,265 contracts  
**Columns:** 170 fields  
**Period:** October 2025 (Period 10)

### Key Metrics
- **Total Revenue:** $893.3M
- **Total Costs:** $637.8M
- **Total Profit:** $393.1M
- **Average Margin:** 31.0%
- **Average Completion:** 78.6%

---

## 🚀 Quick Start

### Run Pre-Built Analytics
```bash
python3 query_wip.py
```

Generates 8 comprehensive business reports:
1. Portfolio Health by Contract Status
2. Margin Distribution Analysis
3. At-Risk Contracts (Open, Low Margin)
4. Regional Performance Comparison
5. Top 15 Customers by Revenue
6. Service Type Performance
7. Top 20 Project Managers
8. Large Projects (>$5M Revenue)

### Custom SQL Queries

**Interactive Mode:**
```bash
python3 custom_query.py
```

Commands: `schema`, `examples`, `quit`

**Single Query:**
```bash
python3 custom_query.py 'SELECT Region, COUNT(*) FROM wip GROUP BY Region'
```

### Python Integration
```python
import duckdb
conn = duckdb.connect('wip_analysis.duckdb', read_only=True)
result = conn.execute("SELECT * FROM wip LIMIT 10").df()
print(result)
conn.close()
```

---

## 📋 Table Schema

**Table Name:** `wip`

### Key Columns

| Column | Type | Description |
|--------|------|-------------|
| `Contract` | VARCHAR | Unique contract identifier |
| `Description` | VARCHAR | Project description |
| `Customer Name` | VARCHAR | Client name |
| `Contract Status` | VARCHAR | Open, Soft-Closed, Hard-Closed, InterCo Elim |
| `Revenue To Date` | DOUBLE | Total revenue recognized ($) |
| `Costs To Date` | DOUBLE | Total costs incurred ($) |
| `Gross Profit` | DOUBLE | Revenue - Costs ($) |
| `Gross Profit %` | DOUBLE | Margin percentage (0-1 scale) |
| `% Complete` | DOUBLE | Project completion (0-1 scale) |
| `Region` | VARCHAR | Geographic region |
| `PM Name` | VARCHAR | Project Manager |
| `ServiceType` | VARCHAR | Soils, Rockfall/Limited Access, etc. |
| `Backlog Revenue` | DOUBLE | Remaining revenue to recognize ($) |
| `WIPMth` | TIMESTAMP | WIP month date |
| `Start Month` | TIMESTAMP | Contract start date |

**Note:** Column names with spaces require double quotes in SQL:
```sql
SELECT "Revenue To Date", "Gross Profit %"  -- ✓ Correct
SELECT Revenue To Date, Gross Profit %      -- ✗ Error
```

---

## 💡 Example Queries

### 1. Contracts by Region
```sql
SELECT Region, 
       COUNT(*) as contracts,
       ROUND(SUM("Revenue To Date")/1000000, 2) as revenue_m
FROM wip
GROUP BY Region
ORDER BY revenue_m DESC;
```

### 2. Open Contracts Near Completion
```sql
SELECT Contract, 
       "Customer Name",
       ROUND("% Complete" * 100, 1) as complete_pct,
       ROUND("Backlog Revenue"/1000, 2) as backlog_k
FROM wip
WHERE "Contract Status" = 'Open'
  AND "% Complete" > 0.90
ORDER BY "Backlog Revenue" DESC
LIMIT 20;
```

### 3. Margin Analysis by Service Type
```sql
SELECT ServiceType,
       COUNT(*) as contracts,
       ROUND(AVG("Gross Profit %") * 100, 1) as avg_margin_pct,
       ROUND(SUM("Gross Profit")/1000000, 2) as total_profit_m
FROM wip
WHERE ServiceType IS NOT NULL
GROUP BY ServiceType
ORDER BY total_profit_m DESC;
```

### 4. PM Performance Dashboard
```sql
SELECT "PM Name",
       COUNT(*) as contracts,
       ROUND(SUM("Revenue To Date")/1000000, 2) as revenue_m,
       ROUND(AVG("Gross Profit %") * 100, 1) as avg_margin_pct,
       ROUND(AVG("% Complete") * 100, 1) as avg_complete_pct
FROM wip
WHERE "PM Name" IS NOT NULL
  AND "Revenue To Date" > 0
GROUP BY "PM Name"
HAVING COUNT(*) >= 5
ORDER BY revenue_m DESC;
```

### 5. Identify Loss-Making Contracts Still Open
```sql
SELECT Contract,
       Description,
       "Customer Name",
       Region,
       "PM Name",
       ROUND("Gross Profit %"  * 100, 1) as margin_pct,
       ROUND("Revenue To Date"/1000000, 2) as revenue_m
FROM wip
WHERE "Contract Status" = 'Open'
  AND "Gross Profit %" < 0
  AND "Revenue To Date" > 100000
ORDER BY "Revenue To Date" DESC;
```

### 6. Backlog by Region
```sql
SELECT Region,
       COUNT(*) as open_contracts,
       ROUND(SUM("Backlog Revenue")/1000000, 2) as backlog_m,
       ROUND(AVG("% Complete") * 100, 1) as avg_complete_pct
FROM wip
WHERE "Contract Status" = 'Open'
  AND Region IS NOT NULL
GROUP BY Region
ORDER BY backlog_m DESC;
```

### 7. Customer Concentration Analysis
```sql
SELECT "Customer Name",
       COUNT(*) as contracts,
       ROUND(SUM("Revenue To Date")/1000000, 2) as revenue_m,
       ROUND(SUM("Revenue To Date") / 
             (SELECT SUM("Revenue To Date") FROM wip) * 100, 2) as pct_of_total
FROM wip
WHERE "Customer Name" IS NOT NULL
GROUP BY "Customer Name"
ORDER BY revenue_m DESC
LIMIT 10;
```

### 8. Completion Status Distribution
```sql
SELECT 
    CASE 
        WHEN "% Complete" < 0.25 THEN '0-25%'
        WHEN "% Complete" < 0.50 THEN '25-50%'
        WHEN "% Complete" < 0.75 THEN '50-75%'
        WHEN "% Complete" < 1.00 THEN '75-99%'
        ELSE '100%'
    END as completion_bucket,
    COUNT(*) as contracts,
    ROUND(SUM("Revenue To Date")/1000000, 2) as revenue_m
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
    END;
```

---

## 🔄 Updating the Database

If the Excel file changes, reload the database:

```bash
python3 setup_duckdb.py
```

This will:
- Drop existing table
- Reload data from Excel
- Recreate the `wip` table
- Verify record count

---

## 📈 Data Quality Notes

### Completeness
- ✅ Contract, Description, Status: 100% complete
- ⚠️ Customer Name: 98.7% complete (30 nulls)
- ⚠️ Region: 98.7% complete (30 nulls)
- ⚠️ PM Name: 95.8% complete (95 nulls)
- ⚠️ ServiceType: 98.9% complete (26 nulls)

### Known Data Patterns
- **Negative Revenue:** 96 records (4.2%) - likely reversals/adjustments
- **Loss-Making Contracts:** 104 records (4.6%) - margins < 0%
- **Zero Revenue:** 336 records (14.8%) - new/inactive contracts
- **InterCo Eliminations:** 163 records - internal transfers

### Recommended Filters

For clean analysis, exclude edge cases:
```sql
WHERE "Contract Status" NOT IN ('InterCo Elim', 'ASC 606 Adjustment')
  AND "Revenue To Date" > 0
  AND Region IS NOT NULL
  AND "PM Name" IS NOT NULL
```

---

## 🎯 Business Insights from Initial Analysis

### Portfolio Health
- **Open Contracts:** 1,217 contracts, $564.6M revenue, 36.3% avg margin
- **Soft-Closed:** 871 contracts, $384.1M revenue, 23.4% avg margin
- **Completion:** Open contracts avg 67% complete, Soft-Closed avg 92%

### Margin Distribution
- **High Margin (>30%):** 1,289 contracts, $602.2M revenue (68% healthy)
- **Low Margin (0-15%):** 169 contracts, $117.7M revenue (need attention)
- **Losses (<0%):** 104 contracts, $50.7M revenue (post-mortem needed)

### Top Regions
1. **Access:** $322.0M revenue, 403 contracts, 14.9% margin, $60.5M backlog
2. **Southeast:** $181.8M revenue, 395 contracts, 30.2% margin, $55.3M backlog
3. **Ohio Valley:** $180.6M revenue, 770 contracts, 36.4% margin, $41.7M backlog

### At-Risk Contracts
- **15 contracts** with >$1M revenue, <15% margin, still open
- **Top concern:** Contract 240713TNCD ($12.8M revenue, 0% margin, 99% complete)

---

## 💼 Next Steps

### For Continuous Use

**In your next conversation, simply say:**
- "Query my WIP DuckDB database"
- "Run a SQL query on WIP analysis"
- "Show me contracts in Ohio Valley region"

No reload needed unless Excel file changes.

### For Deeper Analysis

Consider:
1. **Time Series:** Track margin trends across WIP months
2. **Predictive:** Flag contracts likely to exceed budget
3. **Visualizations:** Create charts with `matplotlib` or `plotly`
4. **Power BI:** Connect directly to DuckDB for dashboards
5. **Automated Reports:** Schedule Python scripts via cron

---

## 🛠️ Tools & Requirements

**Python Packages:**
- `pandas` - Data manipulation
- `openpyxl` - Excel reading
- `duckdb` - SQL database
- `matplotlib`, `seaborn`, `plotly` - Visualizations (optional)

**Installation:**
```bash
pip install pandas openpyxl duckdb matplotlib seaborn plotly
```

---

## 📞 Support

For questions or custom analysis requests, run:
```bash
python3 custom_query.py
```

Then type `examples` to see sample queries.
