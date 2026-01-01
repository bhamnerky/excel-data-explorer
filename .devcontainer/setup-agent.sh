#!/bin/bash
# Setup script for Excel Data Explorer Agent in Codespace

echo "🔧 Setting up Excel Data Explorer Agent..."
echo "=========================================="

# Create agent prompts directory
mkdir -p /workspaces/ExploreData/.vscode/prompts

# Copy agent prompt if not already present
if [ ! -f "/workspaces/ExploreData/.vscode/prompts/DataExplorerAgent.agent.md" ]; then
    echo "📝 Installing agent prompt..."
    
    # Check if running from local or need to download
    if [ -f "/tmp/DataExplorerAgent.agent.md" ]; then
        cp /tmp/DataExplorerAgent.agent.md /workspaces/ExploreData/.vscode/prompts/
        echo "✓ Agent prompt installed"
    else
        echo "⚠️  Agent prompt not found. Please upload manually to:"
        echo "   .vscode/prompts/DataExplorerAgent.agent.md"
    fi
fi

# Create workspace directories
echo "📁 Creating workspace structure..."
mkdir -p /workspaces/ExploreData/ExcelIn
mkdir -p /workspaces/ExploreData/ExcelIn/charts
mkdir -p /workspaces/ExploreData/ExcelIn/summary_charts

# Verify Python packages
echo "🐍 Verifying Python environment..."
python3 -c "import pandas, duckdb, matplotlib, seaborn, plotly" 2>/dev/null
if [ $? -eq 0 ]; then
    echo "✓ All Python packages installed"
else
    echo "⚠️  Some packages missing. Run: pip install -r .devcontainer/requirements-codespace.txt"
fi

# Check VS Code extensions
echo "🔌 VS Code Extensions (will auto-install on first launch):"
echo "   • Python (ms-python.python)"
echo "   • Jupyter (ms-toolsai.jupyter)"
echo "   • GitHub Copilot (GitHub.copilot)"
echo "   • Markdown PDF (yzane.markdown-pdf)"
echo "   • Pylance (ms-python.vscode-pylance)"

echo ""
echo "=========================================="
echo "✅ Setup complete!"
echo ""
echo "📋 Next Steps:"
echo "1. Upload your Excel file to: ExcelIn/"
echo "2. Open GitHub Copilot Chat"
echo "3. Type: @DataExplorerAgent Let's explore [your file.xlsx]"
echo "4. The agent will guide you through analysis, DuckDB setup, and visualization"
echo ""
echo "📖 Agent Capabilities:"
echo "   • Excel structure analysis"
echo "   • DuckDB SQL-queryable database setup"
echo "   • Data quality assessment"
echo "   • Static charts (Matplotlib)"
echo "   • Interactive visualizations (Plotly)"
echo "   • Auto-generate summary documents with charts"
echo "   • PDF export (via Markdown PDF extension)"
echo ""
