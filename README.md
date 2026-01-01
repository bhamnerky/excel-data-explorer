# 📊 Excel Data Explorer - Codespace Ready

This directory is configured as a **GitHub Codespace** for collaborative Excel data analysis using AI.

## 🚀 Quick Start

### For You (Repository Owner)

1. **Copy agent prompt to repository:**
   ```bash
   mkdir -p .vscode/prompts
   cp ~/Library/Application\ Support/Code/User/prompts/DataExplorerAgent.agent.md .vscode/prompts/
   ```

2. **Commit and push:**
   ```bash
   git add .devcontainer/ .vscode/prompts/ CODESPACE_SETUP.md README.md
   git commit -m "Add Excel Data Explorer Codespace configuration"
   git push origin main
   ```

3. **Share repository URL** with your peer

### For Your Peer (No Installation Required)

1. **Open repository on GitHub.com**
2. Click green **"<> Code"** button
3. Select **"Codespaces"** tab
4. Click **"Create codespace on main"**
5. Wait 2-3 minutes for automatic setup
6. Upload Excel file to `ExcelIn/` folder
7. Open Copilot Chat and type: `@DataExplorerAgent Let's explore [your file.xlsx]`

**That's it!** No VS Code, Python, or package installation needed.

---

## 📁 Directory Structure

```
ExploreData/                     ← Root of Codespace
├── .devcontainer/               ← Codespace configuration
│   ├── devcontainer.json        ← Container setup
│   ├── requirements-codespace.txt  ← Python packages
│   └── setup-agent.sh           ← Post-creation script
├── .vscode/
│   └── prompts/
│       └── DataExplorerAgent.agent.md  ← AI agent prompt
├── ExcelIn/                     ← Upload Excel files here
│   ├── charts/                  ← Interactive HTML visualizations
│   ├── summary_charts/          ← PNG charts for embedding
│   ├── *.xlsx                   ← Your Excel files
│   ├── setup_duckdb.py          ← Auto-generated database setup
│   ├── query_*.py               ← Auto-generated query scripts
│   └── *.duckdb                 ← Persistent SQL databases
├── CODESPACE_SETUP.md           ← Full deployment guide
└── README.md                    ← This file
```

---

## 🎯 What the Agent Does

1. **Excel Analysis**: Inspects structure, detects header issues, profiles data
2. **Data Quality**: Assesses completeness, identifies issues, provides recommendations
3. **DuckDB Setup**: Creates persistent SQL-queryable database (97% token savings)
4. **Helper Scripts**: Generates setup, query, and custom query Python scripts
5. **Visualizations**: 
   - Static charts (Matplotlib/Seaborn PNG)
   - Interactive charts (Plotly HTML with hover/zoom/filter)
   - Excel-style slicers (dropdown filters)
6. **Summary Documents**: Auto-generates markdown with embedded charts
7. **PDF Export**: Instructions for creating professional PDF reports

---

## 💰 Cost

**GitHub Codespaces Free Tier:**
- 60 hours/month per user
- 15 GB storage
- **Cost for typical usage: $0/month**

Sufficient for most Excel data exploration work!

---

## 📖 Full Documentation

See [CODESPACE_SETUP.md](CODESPACE_SETUP.md) for complete deployment guide including:
- Detailed configuration
- Security best practices
- Troubleshooting
- Sharing options
- Cost analysis

---

## ✅ Pre-flight Checklist

Before pushing to GitHub:

- [ ] Agent prompt copied to `.vscode/prompts/DataExplorerAgent.agent.md`
- [ ] `.gitignore` updated to exclude sensitive Excel files
- [ ] `CODESPACE_SETUP.md` reviewed
- [ ] Test Codespace created successfully

---

## 🔒 Security Note

Add to `.gitignore` to prevent committing sensitive data:

```gitignore
# Excel data files
*.xlsx
*.xls
*.xlsm

# DuckDB databases
*.duckdb
*.duckdb.wal

# Generated outputs
charts/
summary_charts/
*_Analysis_Summary.md
*_Analysis_Summary.pdf
```

---

**Questions?** See [CODESPACE_SETUP.md](CODESPACE_SETUP.md) for full details.
