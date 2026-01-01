# Quick Start Guide - WIP Analysis

## 🚀 You're All Set!

Your Excel file `casing.xlsx` has been analyzed and loaded into a SQL-queryable database.

---

## 📊 What's Been Created

### 1. Analysis Summary
📄 **[WIP_Analysis_Summary.md](WIP_Analysis_Summary.md)**
- Executive summary with key metrics
- 6 embedded visualizations
- Business insights and recommendations
- At-risk contract identification
- Top performer analysis

### 2. DuckDB Database
🦆 **wip_analysis.duckdb** (2.5MB)
- 2,265 contracts loaded
- 170 columns available
- Persistent SQL database
- Ready for querying

### 3. Analysis Tools

#### Pre-Built Analytics
```bash
python3 query_wip.py
```
Runs 8 comprehensive business reports:
- Portfolio health by status
- Margin distribution
- At-risk contracts
- Regional performance
- Top customers
- Service type analysis
- PM leaderboard
- Large projects

#### Custom SQL Queries
```bash
python3 custom_query.py
```
Interactive mode with commands:
- `schema` - View table structure
- `examples` - Sample queries
- `quit` - Exit
- Or enter any SQL query

#### Chart Generation
```bash
python3 generate_charts.py
```
Creates 6 professional visualizations in `charts/` directory

---

## 💡 Example Queries

### Find High-Margin Contracts in a Region
```bash
python3 custom_query.py 'SELECT Contract, "Customer Name", "Gross Profit %" * 100 as margin FROM wip WHERE Region = '"'"'Access'"'"' AND "Gross Profit %" > 0.40 ORDER BY "Gross Profit %" DESC LIMIT 10'
```

### List Open Contracts by PM
```bash
python3 custom_query.py 'SELECT "PM Name", COUNT(*) as count, SUM("Revenue To Date")/1000000 as revenue_m FROM wip WHERE "Contract Status" = '"'"'Open'"'"' AND "PM Name" IS NOT NULL GROUP BY "PM Name" ORDER BY revenue_m DESC'
```

---

## 📈 Key Findings

### Portfolio Health ✅
- **$893.3M** total revenue
- **44.0%** overall margin
- **68%** of contracts above 30% target

### Areas of Concern ⚠️
- **15 contracts** at-risk (>$1M, <15% margin, open)
- **$51.8M** total at-risk exposure
- **Access region** needs margin improvement (14.9% vs 31% company avg)

### Top Performers 🏆
- **Ohio Valley:** 36.4% avg margin (770 contracts)
- **Carl Sarver:** 44.5% avg margin PM
- **WVDOT:** 45.3% margin across 157 contracts

---

## 🔄 Updating Data

When Excel file changes:
```bash
python3 setup_duckdb.py
```

This reloads the database from the source file.

---

## 📚 Documentation

- **[README_DuckDB.md](README_DuckDB.md)** - Complete reference guide
  - Table schema
  - 15+ example SQL queries
  - Python integration examples
  - Data quality notes
  
- **[WIP_Analysis_Summary.md](WIP_Analysis_Summary.md)** - Executive summary
  - Visual analysis with charts
  - Strategic recommendations
  - At-risk contracts
  - Performance benchmarks

---

## 💬 Next Conversation

The database persists! In your next chat, just say:

- "Query my WIP database"
- "Show me contracts in Southeast region"
- "Find contracts with low margins"
- "Analyze PM performance"

No setup needed - everything's ready to go!

---

## 🛠️ Files Overview

```
excel-data-explorer/
├── FilesIn/
│   └── casing.xlsx              # Source Excel file
├── charts/                       # Generated visualizations (6 PNGs)
├── wip_analysis.duckdb          # Persistent SQL database
├── setup_duckdb.py              # Reload from Excel
├── query_wip.py                 # Pre-built analytics
├── custom_query.py              # Interactive SQL
├── generate_charts.py           # Visualization generation
├── WIP_Analysis_Summary.md      # Executive summary
├── README_DuckDB.md             # Complete reference
└── QUICKSTART.md                # This file
```

---

## 🎯 Common Tasks

### View All Regions
```bash
python3 custom_query.py "SELECT DISTINCT Region FROM wip WHERE Region IS NOT NULL ORDER BY Region"
```

### Export Query to CSV
```bash
python3 custom_query.py "SELECT * FROM wip WHERE Region = 'Access' LIMIT 100"
# Then answer 'y' when prompted to save
```

### Check Database Size
```bash
ls -lh wip_analysis.duckdb
```

### View Sample Data
```bash
python3 custom_query.py "SELECT Contract, Description, \"Customer Name\", \"Revenue To Date\" FROM wip LIMIT 10"
```

---

## 📞 Need Help?

Run any script to see usage examples:
```bash
python3 query_wip.py      # Pre-built reports
python3 custom_query.py   # Interactive mode (type 'examples')
```

---

**Happy Analyzing! 🎉**
