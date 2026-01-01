# 🚀 GitHub Codespace Setup - Excel Data Explorer Agent

This guide explains how to deploy the Excel Data Explorer agent to a GitHub Codespace for collaborative use.

## 📋 What Gets Deployed

✅ **Python Environment**
- Python 3.12
- pandas, duckdb, matplotlib, seaborn, plotly
- openpyxl (Excel reading)
- All visualization dependencies

✅ **VS Code Extensions**
- Python + Pylance
- Jupyter Notebooks
- GitHub Copilot + Chat
- Markdown PDF (for report generation)
- PDF Viewer

✅ **Excel Data Explorer Agent**
- Custom agent prompt with 7-step workflow
- Excel exploration capabilities
- DuckDB persistence
- Visualization generation (static + interactive)
- Summary document generation

✅ **Workspace Structure**
- Pre-configured directories
- Helper scripts ready
- Example analysis preserved

---

## 🎯 Quick Deploy (3 Steps)

### Step 1: Push to GitHub Repository

```bash
# In your local workspace
cd /Users/hamnb/VSCode/BuddySDK/DataGenerationGenAI/ExploreData

# Add Codespace configuration
git add .devcontainer/
git add CODESPACE_SETUP.md

# Copy agent prompt to repository (important!)
mkdir -p .vscode/prompts
# Copy your agent file:
cp ~/Library/Application\ Support/Code/User/prompts/DataExplorerAgent.agent.md .vscode/prompts/

git add .vscode/prompts/DataExplorerAgent.agent.md

git commit -m "Add Codespace configuration for Excel Data Explorer agent"
git push origin main
```

### Step 2: Launch Codespace

**On GitHub.com:**
1. Go to your repository
2. Click green **"<> Code"** button
3. Select **"Codespaces"** tab
4. Click **"Create codespace on main"**
5. Wait 2-3 minutes for environment setup

**Environment will auto-install:**
- All Python packages
- VS Code extensions
- Directory structure

### Step 3: Upload Excel File & Start

1. **Upload your Excel file:**
   - Drag file into `ExcelIn/` folder
   - Or use Upload Files button in Explorer

2. **Open Copilot Chat:**
   - Click chat icon in left sidebar
   - Or press `Cmd/Ctrl + Shift + I`

3. **Invoke agent:**
   ```
   @DataExplorerAgent Let's explore casing.xlsx
   ```

4. **Agent will:**
   - Analyze Excel structure
   - Setup DuckDB database
   - Generate visualizations
   - Create summary document
   - Guide through analysis

---

## 👥 Sharing with Peers

### Option A: Direct Codespace Access (Same Repository)

**For team members with repository access:**
1. They navigate to your GitHub repo
2. Click **"<> Code"** → **"Codespaces"**
3. Create their own Codespace instance
4. Upload their Excel files
5. Use agent independently

**Pros:**
- Each person gets isolated environment
- No conflicts
- Free tier: 60 hours/month per user

### Option B: Share Running Codespace (Live Collaboration)

**For real-time collaboration:**
1. In your Codespace, click **"Share"** icon (top-right)
2. Set permissions (read-only or read-write)
3. Share generated URL
4. Peer can join your session live

**Pros:**
- See each other's work in real-time
- Pair programming style
- Single environment

**Cons:**
- Counts against your hours
- Both need GitHub accounts

### Option C: Export Container (Advanced)

**For offline/air-gapped use:**
1. Build devcontainer locally
2. Export as Docker image
3. Share image file
4. Peer runs in Docker Desktop

---

## 📦 What Your Peer Needs

**Minimum Requirements:**
- ✅ GitHub account (free tier works)
- ✅ Modern web browser
- ✅ Internet connection
- ❌ **NO VS Code installation needed**
- ❌ **NO Python installation needed**
- ❌ **NO package management needed**

**Everything runs in the cloud!**

---

## 🔧 Configuration Details

### Installed Extensions

```json
"extensions": [
  "ms-python.python",           // Python language support
  "ms-toolsai.jupyter",         // Notebook support
  "GitHub.copilot",             // AI coding assistant
  "GitHub.copilot-chat",        // Chat interface
  "yzane.markdown-pdf",         // PDF export from markdown
  "tomoki1207.pdf",             // PDF viewer
  "ms-python.vscode-pylance"    // Python type checking
]
```

### Python Packages

See [.devcontainer/requirements-codespace.txt](.devcontainer/requirements-codespace.txt)

**Key packages:**
- `pandas` - Data manipulation
- `duckdb` - SQL database
- `matplotlib`, `seaborn` - Static charts
- `plotly` - Interactive visualizations
- `openpyxl` - Excel reading

### Directory Structure

```
/workspaces/ExploreData/
├── .devcontainer/
│   ├── devcontainer.json
│   ├── requirements-codespace.txt
│   └── setup-agent.sh
├── .vscode/
│   └── prompts/
│       └── DataExplorerAgent.agent.md
├── ExcelIn/
│   ├── charts/                 (interactive HTML)
│   ├── summary_charts/         (PNG for embedding)
│   ├── [your Excel files]
│   ├── setup_duckdb.py
│   ├── query_*.py
│   └── *.duckdb               (persistent databases)
└── CODESPACE_SETUP.md
```

---

## 🎓 Usage Example

**Your peer's workflow:**

1. **Open Codespace** (from GitHub repo)
   - Automatic setup (2-3 min)
   - All extensions installed
   - Python environment ready

2. **Upload Excel file**
   ```
   ExcelIn/sales_data.xlsx
   ```

3. **Start Copilot Chat**
   ```
   @DataExplorerAgent Let's explore sales_data.xlsx
   ```

4. **Agent responds with:**
   - Dataset overview (rows, columns, date ranges)
   - Data quality assessment
   - Key business metrics
   - Recommendation to setup DuckDB

5. **Continue conversation**
   ```
   @DataExplorerAgent Setup DuckDB for this dataset
   ```

6. **Agent creates:**
   - `setup_duckdb.py` - Database initialization
   - `query_sales.py` - Pre-built business queries
   - `custom_query.py` - Interactive SQL runner
   - `sales_analysis.duckdb` - 2-3 MB persistent database

7. **Generate visualizations**
   ```
   @DataExplorerAgent Generate summary charts
   ```

8. **Agent produces:**
   - 5 PNG charts in `summary_charts/`
   - Updated markdown with embedded images
   - Interactive HTML versions in `charts/`

9. **Create PDF report**
   - Right-click markdown file
   - "Markdown PDF: Export (pdf)"
   - PDF with all charts embedded

10. **Download results**
    - Right-click any file → Download
    - Or download entire folder

---

## 💰 Cost Considerations

**GitHub Codespaces Pricing:**
- **Free Tier**: 60 hours/month + 15 GB storage
- **2-core**: 60 hours free
- **4-core**: 30 hours free

**For Excel Data Explorer:**
- **Recommended**: 2-core (sufficient for most datasets)
- **Storage**: ~500 MB per analysis (Excel + DuckDB + charts)
- **Typical usage**: 5-10 hours/month per user

**Cost if over free tier:**
- 2-core: $0.18/hour
- 4-core: $0.36/hour
- Storage: $0.07/GB/month

**Example monthly cost:**
- User 1 (65 hours): 5 hours × $0.18 = $0.90
- User 2 (40 hours): $0 (within free tier)
- Storage (2 GB): 2 × $0.07 = $0.14
- **Total**: ~$1/month beyond free tier

---

## 🔒 Security & Data

**Data Storage:**
- Excel files stay in Codespace (not in repo by default)
- Add `*.xlsx` and `*.duckdb` to `.gitignore`
- Codespace storage is private to user/workspace

**Best Practices:**
1. Don't commit sensitive Excel files to Git
2. Use `.gitignore` for data files
3. Download results before deleting Codespace
4. Codespaces auto-delete after 30 days inactivity

**`.gitignore` additions:**
```gitignore
# Excel data files
*.xlsx
*.xls
*.xlsm

# DuckDB databases
*.duckdb
*.duckdb.wal

# Generated charts
charts/*.html
charts/*.png
summary_charts/*.png

# Generated summaries (optional)
*_Analysis_Summary.md
*_Analysis_Summary.pdf
```

---

## 🐛 Troubleshooting

### "Agent not found"

**Solution:**
```bash
# Check agent prompt exists
ls -la .vscode/prompts/DataExplorerAgent.agent.md

# If missing, copy from local
cp ~/Library/Application\ Support/Code/User/prompts/DataExplorerAgent.agent.md .vscode/prompts/

# Reload VS Code window
# Cmd/Ctrl + Shift + P → "Developer: Reload Window"
```

### "Package not found"

**Solution:**
```bash
# Reinstall dependencies
pip install -r .devcontainer/requirements-codespace.txt

# Verify installation
python3 -c "import pandas, duckdb, matplotlib, plotly"
```

### "Excel file too large"

**Solution:**
- Codespace upload limit: ~100 MB per file
- For larger files: Use `wget` or `curl` to download from URL
- Or split CSV and import in chunks

### "Out of storage"

**Solution:**
```bash
# Check usage
df -h

# Clean up old databases
rm ExcelIn/*.duckdb

# Remove old charts
rm -rf ExcelIn/charts/*
```

---

## 🚀 Next Steps

1. **Test locally first** (optional):
   ```bash
   # Install devcontainer CLI
   npm install -g @devcontainers/cli
   
   # Build and test
   devcontainer build .
   devcontainer up .
   ```

2. **Commit to Git:**
   ```bash
   git add .devcontainer/ .vscode/prompts/ CODESPACE_SETUP.md
   git commit -m "Add Codespace configuration"
   git push
   ```

3. **Create Codespace** on GitHub.com

4. **Share repository URL** with peer

5. **Peer creates their Codespace** (isolated environment)

---

## 📖 Additional Resources

- **GitHub Codespaces Docs**: https://docs.github.com/codespaces
- **Dev Container Spec**: https://containers.dev/
- **Agent Prompt Guide**: See `.vscode/prompts/DataExplorerAgent.agent.md`
- **Example Analysis**: See `ExcelIn/WIP_Analysis_Summary.md`

---

## ✅ Deployment Checklist

- [ ] `.devcontainer/devcontainer.json` created
- [ ] `.devcontainer/requirements-codespace.txt` created
- [ ] `.devcontainer/setup-agent.sh` created (optional)
- [ ] `.vscode/prompts/DataExplorerAgent.agent.md` copied to repo
- [ ] `.gitignore` updated to exclude data files
- [ ] `CODESPACE_SETUP.md` (this file) added
- [ ] Changes committed to Git
- [ ] Pushed to GitHub
- [ ] Test Codespace created successfully
- [ ] Agent accessible in Copilot Chat
- [ ] Python packages installed correctly
- [ ] Extensions loaded properly
- [ ] Excel file upload works
- [ ] DuckDB setup functional
- [ ] Chart generation works
- [ ] PDF export works
- [ ] Shared repository URL with peer

---

**🎉 Your peer can now use the Excel Data Explorer agent without any local installation!**
