# GitHub Codespaces Cost Analysis
## For Peers Using excel-data-explorer Repository

**Analysis Date:** January 1, 2026  
**Repository:** bhamnerky/excel-data-explorer  
**Use Case:** Excel data analysis with DuckDB and Python

---

## Executive Summary

For a peer to use this repository with their own GitHub account and Codespace:

- **Minimum Annual Cost:** $120/year (Copilot subscription only)
- **Typical Annual Cost:** $120-200/year (Copilot + light Codespace usage)
- **Heavy User Annual Cost:** $200-350/year (daily intensive use)

Most users will stay **within the free Codespace tier** (60 hours/month) and only pay for GitHub Copilot subscription.

---

## 💰 Cost Components

### 1. GitHub Copilot Subscription (REQUIRED)

| Plan | Monthly | Annual | Savings |
|------|---------|--------|---------|
| Individual | $10/month | $100/year | $20 |
| Business | $19/month | $228/year | - |

**Free Options:**
- ✅ Verified students/teachers (GitHub Education)
- ✅ Maintainers of popular open-source projects
- ✅ 30-day free trial (first-time users)

### 2. GitHub Codespaces Usage (PAY-AS-YOU-GO)

#### Free Monthly Allowance

| Resource | Free Tier | Value |
|----------|-----------|-------|
| Compute | 120 core-hours | ~60 hours on 2-core machine |
| Storage | 15 GB | Plenty for this repo (~5 GB typical) |

#### Pricing After Free Tier

**2-core machine** (recommended for this workload):
- **Compute:** $0.18/hour
- **Storage:** $0.07/GB per month

**4-core machine** (unnecessary for Excel analysis):
- **Compute:** $0.36/hour (2× cost)

**8-core machine** (overkill):
- **Compute:** $0.72/hour (4× cost)

---

## 📊 Usage Scenarios & Costs

### Scenario 1: Light User
**Usage:** 10-20 hours/month (occasional analysis)

| Item | Monthly | Annual |
|------|---------|--------|
| Codespaces | $0.00 | $0.00 |
| Copilot | $10.00 | $120.00 |
| **TOTAL** | **$10.00** | **$120.00** |

✅ **Stays within free tier**

---

### Scenario 2: Moderate User
**Usage:** 40-50 hours/month (weekly analysis sessions)

| Item | Monthly | Annual |
|------|---------|--------|
| Codespaces | $0.00 | $0.00 |
| Copilot | $10.00 | $120.00 |
| **TOTAL** | **$10.00** | **$120.00** |

✅ **Still within free tier!**

---

### Scenario 3: Regular User
**Usage:** 80 hours/month (frequent analysis work)

| Item | Calculation | Monthly | Annual |
|------|-------------|---------|--------|
| Codespaces | (80 - 60) × $0.18 | $3.60 | $43.20 |
| Copilot | - | $10.00 | $120.00 |
| **TOTAL** | | **$13.60** | **$163.20** |

⚠️ **20 hours over free tier**

---

### Scenario 4: Heavy User
**Usage:** 100 hours/month (daily data analysis)

| Item | Calculation | Monthly | Annual |
|------|-------------|---------|--------|
| Codespaces | (100 - 60) × $0.18 | $7.20 | $86.40 |
| Copilot | - | $10.00 | $120.00 |
| **TOTAL** | | **$17.20** | **$206.40** |

⚠️ **40 hours over free tier**

---

### Scenario 5: Power User (Daily Work)
**Usage:** 160 hours/month (~8 hours/day × 20 workdays)

| Item | Calculation | Monthly | Annual |
|------|-------------|---------|--------|
| Codespaces | (160 - 60) × $0.18 | $18.00 | $216.00 |
| Copilot | - | $10.00 | $120.00 |
| **TOTAL** | | **$28.00** | **$336.00** |

❌ **100 hours over free tier**

---

## 🎯 Cost Optimization Strategies

### 1. ✅ Always Stop Codespace When Not in Use
- **Auto-stop:** Default 30 minutes of inactivity
- **Manual stop:** Click "Stop Codespace" in VS Code
- **Impact:** Can reduce costs by 50-70%

**Example:**
- Running 8 hours/day but working only 6 hours = 25% waste
- Proper stopping: (120 hrs - 60 free) × $0.18 = **$10.80/mo savings**

### 2. ✅ Use 2-Core Machine (Not 4-Core or 8-Core)
This data analysis workload doesn't need high compute power.

| Machine Type | Cost/Hour | Monthly (100 hrs) | Savings |
|--------------|-----------|-------------------|---------|
| 2-core | $0.18 | $7.20 | Baseline |
| 4-core | $0.36 | $14.40 | +$7.20 waste |
| 8-core | $0.72 | $28.80 | +$21.60 waste |

**Recommendation:** Stick with 2-core unless processing millions of rows.

### 3. ✅ Delete Unused Codespaces
- Storage charges apply to ALL Codespaces (active or stopped)
- Can rebuild from repo anytime
- Delete: GitHub → Codespaces → [···] → Delete

**Impact:**
- 3 old Codespaces @ 5 GB each = 15 GB = **FREE** (within allowance)
- 5 old Codespaces @ 10 GB each = 50 GB = **(50-15) × $0.07 = $2.45/mo waste**

### 4. ✅ Work Locally When Possible
For simple queries or quick checks:
- Clone repo and work locally (free)
- Use Codespace for full analysis sessions
- Hybrid approach saves compute hours

### 5. ⚠️ Avoid Prebuilds (For This Repo)
- Prebuilds speed up startup but consume storage quota
- This repo is lightweight (~2 min cold start)
- **Not worth the storage cost**

---

## 🧮 Total Realistic Cost Estimate

### For Typical Excel Data Analysis Work:

**Most Likely Annual Cost:** $120-180/year

**Breakdown:**
- GitHub Copilot: $120/year (required)
- Codespaces: $0-60/year (depending on usage patterns)
- **Assumes:** 50-70 hours/month with good stop habits

### Best Case:
**$100-120/year** (Copilot only, stays in free tier)

### Worst Case:
**$300-400/year** (heavy daily use without optimization)

---

## 🆓 Free Trial Option

Your peer can **test this setup for FREE** before committing:

1. **Fork the repository** (free)
2. **Sign up for GitHub Copilot** (30-day free trial)
3. **Create Codespace** (free within 120 core-hours/month)
4. **Try for 1 month** = **$0 cost**

After trial:
- If valuable → Subscribe to Copilot ($10/mo)
- If not → Cancel, no cost

---

## 📋 Setup Checklist for Peer

- [ ] Create GitHub account (free)
- [ ] Subscribe to GitHub Copilot ($10/mo or free trial)
- [ ] Fork or clone `bhamnerky/excel-data-explorer` repo
- [ ] Create Codespace (use 2-core machine)
- [ ] Configure auto-stop to 30 minutes
- [ ] Upload Excel files to `FilesIn/` folder
- [ ] Run analysis scripts from `PythonScripts/`
- [ ] Remember to **Stop Codespace** when done!

---

## ❓ FAQ

### Q: Can multiple people share one GitHub account to save costs?
**A:** Technically possible but **NOT recommended**:
- Violates GitHub Terms of Service
- Security risks (shared credentials)
- Conflicts (overwriting each other's work)
- Better: Each person pays $10-15/month for their own

### Q: Is there a team/organization discount?
**A:** Yes! GitHub Copilot Business is $19/user/month but includes:
- Centralized billing
- Admin controls
- License management
- Worth it for teams of 5+

### Q: What if we only use this once a month?
**A:** Consider **local setup** instead:
- Install Python + DuckDB locally (free)
- Use GitHub Copilot in VS Code Desktop ($10/mo)
- No Codespace costs
- **Total: $120/year** (vs $120-200 with Codespace)

### Q: Can we use GitHub Codespaces without Copilot?
**A:** Yes, but you'll lose:
- AI-assisted code generation
- Smart query suggestions
- Documentation help
- **Not worth it** for this AI-powered workflow

### Q: What about data privacy with Copilot?
**A:** GitHub Copilot for Individuals:
- Code snippets sent to OpenAI for suggestions
- Not used for model training (as of 2023+)
- Review GitHub's privacy policy
- For sensitive data: Use Copilot Business (excludes data from training)

---

## 📞 Support & Questions

- **Repository Issues:** https://github.com/bhamnerky/excel-data-explorer/issues
- **GitHub Codespaces Docs:** https://docs.github.com/en/codespaces
- **GitHub Copilot Docs:** https://docs.github.com/en/copilot
- **Pricing Details:** https://github.com/pricing

---

## ✅ Conclusion

**Bottom Line for Your Peer:**

For typical Excel data analysis work using this repository:
- **Expect to pay:** $120-200/year
- **Mostly:** GitHub Copilot subscription ($120/yr)
- **Codespaces:** Likely free or minimal ($0-80/yr)

**The value proposition is strong:**
- AI-assisted data analysis (faster, smarter)
- Zero local setup (works from any device)
- Persistent environment (pick up where you left off)
- Professional DuckDB + Python stack (pre-configured)

**Recommendation:** Start with the free trial, then commit if valuable. For most analysts, **$10-15/month is a bargain** for the productivity gains.

---

*Last Updated: January 1, 2026*
