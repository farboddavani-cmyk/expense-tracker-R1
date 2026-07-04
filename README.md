# 🧾 Ledger Pro — California Business Finance Tracker

> **Designed specifically for California freelancers, sole proprietors, single-member LLCs, and small businesses operating under California Franchise Tax Board (FTB) and IRS requirements.**
>
> Track expenses and income aligned to **Schedule C**, **CA Schedule CA (540)**, and **FTB Publication 984**. Google Sheets sync, receipt capture, mileage logging, home office tracking, CSV export, 6 languages.

**Live App:** https://farboddavani-cmyk.github.io/expense-tracker/
**Play Store:** `com.farboddavani.ledger`
**Version:** 1.0.0

---

## ⚖️ California Tax Compliance — What This App Covers

Ledger Pro is built around the specific requirements of the **California Franchise Tax Board (FTB)** and **IRS Schedule C** for sole proprietors, freelancers, and single-member LLCs doing business in California.

### ✅ California-Specific Features

| Feature | CA/FTB Requirement Addressed |
|---|---|
| **Schedule C Line Categories** | Every expense maps to a Schedule C line (Ln 8–30) per IRS & FTB Pub. 984 |
| **Mileage Logging** | IRS/CA standard rate: $0.70/mi (2025), $0.725/mi (2026) — auto-calculates deduction |
| **Home Office Tracker** | Simplified method: $5/sq ft up to 300 sq ft ($1,500 max) — CA FTB conforms to IRS |
| **Meals (50% deductible)** | Separate category for 50%-deductible meals per CA Schedule CA (540) instructions |
| **Tax Deductible Flag** | Per FTB Pub. 984 — marks ordinary & necessary business expenses |
| **Quarterly Estimated Tax Banner** | CA FTB 540-ES schedule: **30% Q1 / 40% Q2 / 0% Q3 / 30% Q4** (differs from federal!) |
| **CA Estimated Tax Estimator** | Estimates CA state tax at ~9.3% blended rate on net self-employment income |
| **Receipt Photo Capture** | IRS/FTB require adequate records: date, amount, place, business purpose |
| **Google Sheets Export** | Audit-ready records with Schedule C line references, mileage, home office sq ft |
| **Invoice Status Tracking** | Tracks Paid/Unpaid/Overdue/Partial — required for CA income reporting |

### 📋 Expense Categories (Schedule C Aligned)

All categories map directly to IRS Schedule C / CA FTB equivalent lines:

| App Category | Schedule C Line | CA FTB Notes |
|---|---|---|
| Advertising & Marketing | Line 8 | Fully deductible |
| Car & Truck (Mileage) | Line 9 | $0.70/mi (2025) · $0.725/mi (2026) · mileage log required |
| Commissions & Fees | Line 10 | 1099-NEC required if $600+ to individual |
| Contract Labor (1099) | Line 11 | Issue 1099-NEC by Jan 31 for $600+ paid |
| Depreciation & Section 179 | Line 13 | ⚠ CA does NOT conform to federal bonus depreciation (OBBBA 2025) |
| Employee Benefits | Line 14 | Health insurance, retirement plans |
| Home Office | Line 30 | Simplified: $5/sq ft, max 300 sq ft ($1,500) |
| Insurance (Business) | Line 15 | Fully deductible business insurance only |
| Interest (Business Loans) | Line 16 | CA does not cap at 30% ATI like federal |
| Legal & Professional Services | Line 17 | CPA, attorney, consulting fees |
| Meals & Entertainment (50%) | Line 24b | Only 50% deductible — CA follows federal limit |
| Office Supplies & Postage | Line 18 | Fully deductible |
| Rent & Lease | Line 20a/20b | Equipment or business property |
| Repairs & Maintenance | Line 21 | Must be repair not improvement |
| Software & Subscriptions | Line 27a | SaaS, apps, tools |
| Taxes & Licenses | Line 23 | CA LLC Annual Fee ($800+), business licenses |
| Travel (Away from Home) | Line 24a | Must be away from tax home overnight |
| Utilities | Line 25 | Business-use percentage only |
| Wages & Salaries | Line 26 | W-2 employees only |
| Other Business Expense | Line 27a | Catch-all for ordinary & necessary |

### 📅 California Quarterly Estimated Tax Schedule

> ⚠️ **California's schedule differs from the federal IRS schedule — this is one of the most commonly missed CA requirements.**

| Quarter | CA FTB Due Date | % of Annual Estimated Tax | Federal IRS |
|---|---|---|---|
| Q1 | April 15 | **30%** | 25% |
| Q2 | June 15 | **40%** | 25% |
| Q3 | September 15 | **0%** | 25% |
| Q4 | January 15 | **30%** | 25% |

Pay via **FTB Form 540-ES** at [ftb.ca.gov](https://www.ftb.ca.gov). Required if you expect to owe **$500 or more** in CA taxes (vs. $1,000 federal threshold).

### 🏛️ Key California Tax Rules for Freelancers & LLCs

**California Franchise Tax Board (FTB) Requirements:**
- **LLC Annual Fee:** Minimum $800/year even with zero income (Form 3522)
- **CA SDI Rate:** 1.2% of wages (up to ~$145,600 in 2025) — administered by EDD
- **CA Income Tax Rates:** 1%–13.3% progressive (10 brackets) + 1% Mental Health Services Tax on income over $1M
- **CA 2% Rule:** California still applies the 2% AGI floor on miscellaneous itemized deductions (unlike federal post-TCJA)
- **Bonus Depreciation:** CA does NOT conform to the federal OBBBA 2025 bonus depreciation restoration — separate CA adjustment required on Schedule CA (540)
- **Net Operating Loss Suspension:** CA has suspended NOL carryover deductions for tax years 2024–2026 (taxpayers with net business income under $1M are exempt)
- **Home Office:** CA conforms to federal simplified method ($5/sq ft, max $1,500/year)
- **Meals:** CA follows federal 50% deductibility limit for business meals
- **Business Interest:** CA does NOT cap business interest deduction at 30% of ATI (unlike federal)
- **1099-NEC:** Issue to any individual/unincorporated business paid $600+ for services by Jan 31

**Self-Employment Tax (Federal, applies to CA filers):**
- 15.3% on 92.35% of net earnings (12.4% Social Security + 2.9% Medicare)
- Social Security wage base: $176,100 (2026)
- Deduct half of SE tax on Form 1040 (reduces CA taxable income too)
- QBI Deduction: Up to 20% of qualified business income (Section 199A) — federal only, CA does not conform

### 📂 What Records to Keep (FTB Audit Requirements)

Per **FTB Publication 984** and IRS requirements, keep for **7 years**:
- ✅ Receipts for all business expenses (date, amount, vendor, business purpose)
- ✅ Mileage log (date, destination, business purpose, miles — not reconstructed at year-end)
- ✅ Bank and credit card statements
- ✅ Invoices issued and received
- ✅ Home office measurements and photos (for home office deduction)
- ✅ 1099s received (reconcile with your own income records)
- ✅ Contracts and client agreements
- ✅ Quarterly estimated tax payment confirmations (FTB 540-ES)

---

## 📁 Project Structure

```
ledger-pro/
├── index.html              ← Complete single-file PWA
├── manifest.json           ← PWA manifest
├── sw.js                   ← Service worker (offline caching)
├── Code.gs                 ← Google Apps Script backend (14-col expense log)
├── icons/
│   ├── icon-192.png
│   ├── icon-512.png
│   ├── icon-512-maskable.png
│   └── feature-graphic.png
├── PRIVACY_POLICY.md
├── PLAY_STORE_LISTING.md
└── README.md               ← This file
```

---

## 🚀 Deploy to GitHub Pages

### Option 1: GitHub Actions (Recommended)

1. Push this repo to `github.com/farboddavani-cmyk/expense-tracker`
2. Go to **Settings → Pages → Source → GitHub Actions**
3. Push any commit to `main` — it deploys automatically
4. Live at: `https://farboddavani-cmyk.github.io/expense-tracker/`

### Option 2: Manual

```bash
git init
git add .
git commit -m "Initial release"
git remote add origin https://github.com/farboddavani-cmyk/expense-tracker.git
git push -u origin main
# Then enable Pages in repo Settings → Pages → Deploy from branch: main / root
```

---

## 🔧 Google Apps Script Setup

### First-time setup (new spreadsheet)

1. Open Google Sheets → **Extensions → Apps Script**
2. Select all in `Code.gs`, delete, paste the new `Code.gs`
3. Click **Save** (💾)
4. Run `setupSheets()` once from the editor to create tabs with CA-aligned headers
5. **Deploy → New Deployment → Web App**
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Copy the deployment URL → Ledger Pro app → Settings → Apps Script URL

### Updating existing deployment (keep your data)

1. Paste new `Code.gs`, click Save
2. **Deploy → Manage Deployments → Edit → New Version → Deploy**
3. URL stays the same — data is untouched

### Spreadsheet Structure

**Expense Log** (14 columns, Schedule C aligned):
| Date | Vendor | Description | Category | Amount | Currency | Payment Method | Tax Deductible | Schedule C Line | Miles | Sq Ft | Notes | Receipt Link | ID |

**Income Log** (9 columns):
| Date | Client | Invoice # | Amount | Currency | Status | Notes | Invoice Link | ID |

**Dashboard** — Live KPIs: Total Income, Net Position, Total Expenses, Invoices Paid, Unpaid Amount, Tax Deductible, Mileage Summary, CA compliance notes

**Tax Summary** — Accountant-ready: deductible vs non-deductible by Schedule C category, income/expense net, mileage deduction calculation

---

## 🌍 Languages Supported

| Code | Language |
|------|----------|
| `en` | English |
| `es` | Español |
| `fr` | Français |
| `fa` | فارسی (RTL) |
| `de` | Deutsch |
| `pt` | Português |

---

## 💳 In-App Purchase

| Property | Value |
|----------|-------|
| Product ID | `ledger_unlock_full` |
| Type | One-time (INAPP) |
| Price | $9.99 |
| Trial | 10 days from first launch |
| Billing Library | Google Play Billing 7.0.0 |

---

## 🏗️ Tech Stack

- **Frontend:** Single-file HTML + CSS + Vanilla JS
- **PWA:** Web App Manifest + Service Worker
- **Backend:** Google Apps Script (your own deployment)
- **Storage:** Google Sheets (cloud) + localStorage (local)
- **Files:** Google Drive (receipts & invoices)
- **Android:** PWABuilder + Custom Kotlin (TrialManager + BillingManager)
- **Billing:** Google Play Billing Library 7.0.0

---

## ⚠️ Disclaimer

Ledger Pro is a record-keeping and expense tracking tool. It is **not tax advice**. California tax law is complex and changes frequently. Always consult a licensed California CPA or tax professional (enrolled agent, FTB-registered) before filing. The quarterly tax estimates provided in the app are approximations only and do not account for your full tax situation.

---

## 📄 License

Copyright © 2025 Farbod Davani. All rights reserved.

---

*Built with Ledger Pro v1.0.0 — Designed for California FTB & IRS Schedule C compliance*
