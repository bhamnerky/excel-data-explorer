# Excel Data Explorer - Project Overview

An intelligent workspace for analyzing Excel datasets with DuckDB and generating business insights.

## 📂 Project Structure

```
/workspaces/excel-data-explorer/
├── FilesIn/              ← Place your Excel files here
│   └── casing.xlsx       ← Example WIP analysis data
│
├── PythonScripts/        ← All analysis scripts
│   ├── setup_duckdb.py       ← Load Excel into DuckDB
│   ├── query_wip.py          ← Pre-built business queries
│   ├── custom_query.py       ← Interactive SQL querying
│   ├── generate_charts.py    ← Create visualizations
│   └── analyze_casing.py     ← Quick data overview
│
├── AnalysisOut/          ← Generated analysis reports
│   └── WIP_Analysis_Summary.md ← Full analysis with insights
│
├── SetupAndHelp/         ← Documentation and guides
│   ├── README.md             ← This file
│   ├── QUICKSTART.md         ← Quick start guide
│   ├── README_DuckDB.md      ← DuckDB setup guide
│   └── CODESPACE_SETUP.md    ← Codespace configuration
│
├── charts/               ← Generated PNG visualizations
│   ├── 01_regional_revenue.png
│   ├── 02_margin_distribution.png
│   └── ...
│
├── wip_analysis.duckdb   ← Persistent SQL database
│
└── .github/agents/       ← Custom Copilot agent
    └── ExcelAnalyzer.agent.md
```

## 🚀 Quick Start

### Analyze a New Excel File

1. **Place your Excel file** in `FilesIn/`

2. **Load into DuckDB:**
   ```bash
   python3 PythonScripts/setup_duckdb.py
   ```

3. **Run pre-built analytics:**
   ```bash
   python3 PythonScripts/query_wip.py
   ```

4. **Generate charts:**
   ```bash
   python3 PythonScripts/generate_charts.py
   ```

5. **View results** in `AnalysisOut/*_Summary.md`

### Use the Custom Agent

If you have GitHub Copilot, use the **Excel Data Explorer** agent:
- Automatically analyzes Excel files
- Creates DuckDB databases
- Generates business insights
- Produces visualizations

## 🎯 Features

- ✅ **Automated Excel Analysis** - Header detection, data profiling, quality checks
- ✅ **SQL Querying** - Convert Excel to DuckDB for powerful SQL analysis
- ✅ **Business Insights** - Pre-built queries for common business questions
- ✅ **Visualizations** - Professional charts with matplotlib/seaborn
- ✅ **Persistent Storage** - DuckDB databases persist between sessions
- ✅ **Custom Agent** - AI-powered analysis with GitHub Copilot

## 📊 Example Analysis

The workspace includes a complete WIP (Work in Progress) analysis:
- **Dataset:** 2,265 contracts, $893.3M revenue
- **Insights:** Margin distribution, regional performance, at-risk contracts
- **Visualizations:** 6 professional charts
- **Report:** `AnalysisOut/WIP_Analysis_Summary.md`

## 🤝 Sharing This Workspace

### Option 1: Share Repository
1. Add collaborators to this GitHub repo
2. They create their own codespace
3. All files, scripts, and data are included
4. Custom agent available if they have Copilot

### Option 2: Manual Use (No Copilot Needed)
Users can run all scripts manually:
```bash
python3 PythonScripts/setup_duckdb.py
python3 PythonScripts/query_wip.py
python3 PythonScripts/custom_query.py
```

## 📚 Documentation

- **Quick Start:** `SetupAndHelp/QUICKSTART.md`
- **DuckDB Guide:** `SetupAndHelp/README_DuckDB.md`
- **Codespace Setup:** `SetupAndHelp/CODESPACE_SETUP.md`
- **Agent Instructions:** `.github/agents/ExcelAnalyzer.agent.md`

## 🔧 Requirements

- Python 3.x
- pandas, openpyxl, duckdb
- matplotlib, seaborn (for charts)
- GitHub Copilot (optional, for AI agent)

## 💡 Tips

- **DuckDB persists:** No need to reload unless Excel data changes
- **Scripts use relative paths:** Run from project root
- **Charts auto-generate:** Embedded in summary markdown
- **Custom queries:** Use `custom_query.py` for ad-hoc SQL

---

**Ready to analyze your data!** 🎉
