# Job Board Usage Dashboard
**HonorVet Technologies — Weekly Job Board Reporting**

A fully client-side, GitHub Pages-compatible dashboard for analyzing job posting and resume search activity across all job boards. Upload your Excel files each week to instantly refresh all KPIs, charts, and insights — no backend, no server, no configuration.

---

## 🚀 Deploy to GitHub Pages (5 minutes)

### Step 1 — Create the repository
1. Go to [github.com/new](https://github.com/new)
2. Name it something like `job-board-dashboard`
3. Set it to **Private** (recommended for internal data)
4. Click **Create repository**

### Step 2 — Upload the dashboard files
```bash
# Clone, add files, push
git clone https://github.com/YOUR_ORG/job-board-dashboard.git
cd job-board-dashboard

# Copy these files into the folder:
#   index.html
#   style.css
#   app.js
#   assets/   (folder)
#   README.md

git add .
git commit -m "Initial dashboard deployment"
git push origin main
```

### Step 3 — Enable GitHub Pages
1. In the repo, go to **Settings → Pages**
2. Under **Source**, select **Deploy from a branch**
3. Choose branch: `main`, folder: `/ (root)`
4. Click **Save**
5. Your dashboard is live at: `https://YOUR_ORG.github.io/job-board-dashboard/`

> ⏱️ First deployment takes ~1 minute to appear.

---

## 📁 Folder Structure

```
job-board-dashboard/
├── index.html          ← Main dashboard page
├── style.css           ← All styling (no external CSS deps)
├── app.js              ← All logic: parsers, charts, filters, exports
├── assets/             ← Static assets (empty, reserved for logos)
└── README.md           ← This file
```

---

## 📋 How to Use (Weekly)

1. Open your dashboard URL (GitHub Pages or local)
2. On the upload screen, drop or browse:
   - **Job Postings.xlsx** — into the left panel
   - **Resume Searches.xlsx** — into the right panel
3. Click **Generate Dashboard**
4. All KPIs, charts, tables, and insights auto-generate from your files
5. Use the **filters** (date range, board, vertical, recruiter) to drill down
6. Use the **Export** menu to download PDF reports or CSV data

> ✅ No data is ever uploaded to any server. All processing happens in your browser.

---

## 📊 What's Analyzed

### Job Postings File (`Job Postings.xlsx`)
| Sheet | Board | Key Metrics |
|-------|-------|-------------|
| `Dice_Posting` | Dice | Jobs, Views, Applications, Alerts, Status, Vertical |
| `LinkedIn_Posting` | LinkedIn | Jobs, Views, Apply Clicks, Unique Viewers, Recruiter |
| `Indeed Analysis` | Indeed | Jobs, Impressions, Clicks, Applies, Spend, CPA |
| `Website_Posting` | HonorVet Website | Jobs, Applicants, Recruiter, Vertical |
| `DocCafe Data` | DocCafe | Jobs, Views, Applications (physician jobs) |
| `Vivian_Posting` | Vivian | Jobs, Inbound/Proposal candidates, Recruiter |

### Resume Searches File (`Resume Searches.xlsx`)
| Sheet | Board/Source | Key Metrics |
|-------|-------------|-------------|
| `JobDiva Resumes` | All (Indeed, LinkedIn, Monster, CareerBuilder…) | Individual resume views by date/recruiter |
| `Linkedn Analysis` | LinkedIn | InMail sends, responses, accepts, response rate |
| `Indeed Analysis` | Indeed | Smart sourcing: searches, views, contacts used |
| `Monster Analysis` | Monster | Credits used/available, integrated & direct fetches |
| `Resume-Library Analysis` | Resume Library | Monthly searches, views, unlocks by recruiter |
| `Signal Hire` | Signal Hire | Contact unlocks per account |
| `Vivian Candidates` | Vivian | Premium/standard credits, candidates by recruiter |
| `Dice Resume Analysis` | Dice | Resume searches by recruiter (pivot) |
| `Vivian Candidates` | Vivian | Candidate pipeline with discipline/specialty |

---

## 📈 Dashboard Sections

| Section | Contents |
|---------|----------|
| **Overview** | 8 KPI cards (jobs, views, applications, searches, top board, growth %, active recruiters) |
| **Job Postings** | Bar chart, donut share chart, trend line, views vs applications comparison |
| **Resume Activity** | Search/view bar chart, share donut, trend line, LinkedIn InMail chart |
| **Board Details** | Full performance table for postings + searches per board |
| **Recruiters** | Recruiter posting bar, resume search bar, full recruiter summary table |
| **Insights** | 8–12 auto-generated insights: top board, underperformers, spend efficiency, pipeline health |

---

## 🔧 Filters

| Filter | Applies To |
|--------|-----------|
| **Date Range** | All posting and search records with dates |
| **Board** | Posting and search records for selected board |
| **Vertical** | Posting records (IT, Healthcare, BFSI, Pharma, Allied, etc.) |
| **Recruiter** | Both posting and search records |

Resetting filters returns to all data.

---

## 📤 Export Options

| Export | Format | Content |
|--------|--------|---------|
| PDF Report | `.pdf` | Board summaries (postings + searches) |
| Postings CSV | `.csv` | Every filtered posting record |
| Searches CSV | `.csv` | Every filtered search record |
| Recruiter CSV | `.csv` | Recruiter performance summary |
| Chart Images | `.png` | Any chart via ⬇ button on chart header |
| Print View | Browser print | Print-optimized layout |

---

## 🛠 Technical Notes

- **Libraries used** (loaded from CDN, no install needed):
  - [SheetJS (xlsx 0.18.5)](https://sheetjs.com/) — Excel parsing
  - [Chart.js 4.4](https://www.chartjs.org/) — Charts
  - [jsPDF 2.5](https://artskydj.github.io/jsPDF/) — PDF export
  - [html2canvas 1.4](https://html2canvas.hertzen.com/) — PDF rendering
  - [Google Fonts — Inter](https://fonts.google.com/specimen/Inter) — Typography

- **Browser compatibility**: Chrome, Edge, Firefox, Safari (all modern)
- **File size limit**: Excel files up to ~50 MB work without issues
- **Mobile**: Responsive layout; filters and tables scroll horizontally on small screens
- **Offline**: Works without internet once the page is loaded (except Google Fonts)

---

## 🔄 Adding New Boards in the Future

The dashboard auto-detects any new sheets matching the known column patterns. For a completely new board format:
1. Open `app.js`
2. Add a new parser function following the pattern of `parseDicePosting()` etc.
3. Add the parser call inside `parseAllPostings()` or `parseAllSearches()`
4. Optionally add the board color in the `BOARD_CFG` object at the top

---

*HonorVet Technologies · Confidential Internal Report Tool*
