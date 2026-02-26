# Time Exception Slip Automation Tool

> This tool was built with one goal — to make sure our employees get paid as accurately and efficiently as possible. That matters to us. Getting payroll right isn't just a task, it's a responsibility, and this tool helps us deliver on that every single pay period.

Every pay period we have to fill out Time Exception Slips by hand for ~128 employees — writing in names, employee numbers, dates, overtime hours, the whole thing. It is time consuming and honestly, we are human, we are bound to make a few errors here and there. This tool does all of that for us automatically with increased accuracy.

## What it does

### Blank Slip Binder
Upload the employee list, pick the pay period date, and it spits out one big PDF with everyone's slip already filled in — names, employee numbers, dept code, dates, all of it. Sorted alphabetically, ready to print.

### Overtime Entry + Export
Search for an employee, punch in their OT hours per day, and it fills out their slip and builds an Excel sheet for Finance. No more handwriting OT on slips and then retyping it into a spreadsheet.

- **Searchable employee selector** — just type a name and hit Enter
- **Biweekly structure** — Week 1 and Week 2 dates are figured out automatically from the pay period end date
- **OT categories** — OT 1.0, OT 1.5, CTE 1.0, CTE 1.5
- **Live summary table** — see everything laid out before you generate
- **One-click export** — get your filled PDF slips + formatted Excel in one shot

## Time & accuracy improvements

| Task | By hand | With this tool |
|------|---------|----------------|
| Fill ~128 blank slips | ~2 hours | ~10 seconds |
| Enter OT for ~30 employees | ~1.5 hours | ~20 minutes |
| Build the Excel summary | ~30 minutes | Instant |
| Merge everything into one PDF | ~15 minutes | Instant |
| Name/ID/date typos | Common | Eliminated |
| OT math errors | Occasional | Eliminated |
| PDF ↔ Excel mismatches | Possible | Eliminated |
| **Total per pay period** | **~4+ hours** | **~20 minutes** |

That's roughly **100+ hours saved per year** with significantly higher accuracy across the board.

## How to use it

1. Upload your employee CSV (just needs LastName, FirstName, EmployeeNumber)
2. Pick the pay period ending date
3. Use *Blank Slips* for the full binder, or *Overtime Entry* to log OT and get PDFs + Excel
4. Download, print, done

Your data stays in your browser only — nothing gets saved to a server. Hit "Clear & Start New Period" when you're done with a pay cycle.

## UI Themes

Three built-in themes (your pick saves automatically):
- **Default** — clean blue
- **Lissa Mode** — purple & pink
- **Dark** — easy on the eyes

## Quick Start

```bash
pip install -r requirements.txt
python app.py
# → Open http://localhost:5050
```

## Deploy with Docker

```bash
docker build -t ot-slips .
docker run -p 8080:8080 ot-slips
```

Or connect your GitHub repo to [Railway](https://railway.app) — it picks up the Dockerfile and just works.

## Tech Stack

- **Backend**: Python / Flask
- **PDF Generation**: ReportLab (text overlay) + PyPDF2 (merging)
- **Excel Export**: openpyxl
- **Frontend**: Vanilla HTML/CSS/JS — no build tools, no frameworks
- **Data**: Browser localStorage (no database needed)
- **Deploy**: Docker / Gunicorn

## Project Structure

```
├── app.py                             # Backend + PDF/Excel generation
├── templates/index.html               # The web UI
├── static/
│   ├── style.css                      # Styles (3 themes)
│   └── app.js                         # Frontend logic + state
├── OT_Time_Exception_Slip_Sample.pdf  # PDF template (Form P/R-107)
├── requirements.txt
├── Dockerfile
└── README.md
```

## Author

Built by [David Godinez](https://github.com/davidgo24) — custom internal gov tool for municipal payroll.
