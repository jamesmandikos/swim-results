# Swimming Results — Working Guide

## Folder structure

```
Swimming Results/
├── Scripts/
│   ├── swimming_results.py        — main script (PBs + HTML reports)
│   └── update_qualifying_times.py — annual QT update script
├── Margot/
│   ├── MargotSwimTimes.xlsx       — master spreadsheet
│   ├── margot_club_rankings.html  — club comparison report (open in browser)
│   ├── margot_swimmers.json       — Brompton SC female 2014 swimmer list
│   └── margot_previous_bests.json — saved times for progress arrows
├── Ava/
│   ├── AvaSwimTimes.xlsx          — Ava's spreadsheet
│   ├── ava_club_rankings.html     — club comparison report (open in browser)
│   └── ava_swimmers.json          — Chelsea & Westminster female 2017 swimmer list
├── Reference/
│   ├── middlesex_county_qt_2026.pdf
│   └── london_regional_qt_2026.pdf
└── README.md
```

---

## Task 1 — Update after every gala

Type `/MargotSwimming` in Claude Code. It will run automatically and give a plain-English summary of what changed.

Or run directly from Terminal:
```bash
cd "/Users/jamesmandikos/Documents/Claude Cowork/Outputs/Swimming Results/Scripts"
python3 swimming_results.py
```

**What it updates:**
- `MargotSwimTimes.xlsx` → PBs updated where the site has a faster time
- `AvaSwimTimes.xlsx` → same for Ava
- `margot_club_rankings.html` → rebuilt with latest times and club rankings
- `ava_club_rankings.html` → rebuilt for Ava's club

Takes under 2 minutes.

---

## Task 2 — Add a new swimmer to the comparison

1. Go to swimmingresults.org/individualbest/ and search by surname
2. Find the entry matching the club and year of birth — note the member ID number
3. Edit the relevant JSON file:
   - Margot's club: `Margot/margot_swimmers.json` (Brompton SC, female, born 2014)
   - Ava's club: `Ava/ava_swimmers.json` (Chelsea & Westminster, female, born 2017)
4. Add: `{ "name": "First Last", "member_id": "XXXXXXX" }`
5. Run `/MargotSwimming`

---

## Task 3 — Update qualifying times (once a year, October/November)

Download the two official PDFs for the new season:
- **Middlesex County:** harrowswim.com — search "Middlesex County QT [year]" → save to `Reference/middlesex_county_qt_[year].pdf`
- **London Regional:** swimming.org/library — search "SE London Summer Championships QT [year]" → save to `Reference/london_regional_qt_[year].pdf`

Then type `/UpdateQualifyingTimes` in Claude Code. It will parse the PDFs and update all QT columns.

**Do not use swimmingresults.org/qt/ for county or regional times** — those values differ from the official published PDFs.

---

## Task 4 — Start of a new season

Tell Claude Code and it will:
- Create a new `20XX-Times` tab in both spreadsheets (copying the previous year's structure)
- Update the QT columns with the new age group (Margot moves from 12→13→14 etc., Ava similarly)
- Update year references in both scripts

Margot's `2027-Times` tab already exists — it shows the age-13 county and regional qualifying times she needs to target next season.

---

## Understanding the HTML reports

Open `margot_club_rankings.html` or `ava_club_rankings.html` in any browser.

**Summary panel** (top): coloured pills showing events where she currently holds Q (qualifies), C (consideration), or is 1st in the club.

**Table**: one row per swimmer, columns grouped by stroke. Each cell shows time, club rank, and date.
- Green cell = 1st in club
- Amber = 2nd, Orange = 3rd
- Yellow row = Margot / Ava
- ▲ marker = new PB since the last run (click "Hide markers" to remove)

**Sorting**: click any column header to sort. ↑ = ascending, ↓ = descending.

**Filtering**: use the "Filter columns" panel to show/hide specific strokes or column types (SC / LC / LC→SC converted).

**History popup**: click any of Margot's (or Ava's) time cells to see a chart of all their recorded swims for that event, with County QT, County Cons, Regional Cons and Regional QT reference lines. The table below shows each swim with the qualifying standards alongside for comparison.

---

## Spreadsheet column layout (2026-Times and 2027-Times)

| Col | Content |
|-----|---------|
| B | Course (S = Short Course, L = Long Course) |
| C | Event |
| D | PB |
| E | Date obtained |
| F | National ranking |
| G | County QT |
| H | County QT difference vs PB |
| I | County Cons (consideration time) |
| J | County Cons difference vs PB |
| K | County Qual — **Q** = qualifies, **C** = consideration, blank = neither |
| L | County ranking |
| M | Regional QT |
| N | Regional QT difference vs PB |
| O | Regional Cons |
| P | Regional Cons difference vs PB |
| Q | Regional Qual — **Q** / **C** / blank |
| R | Regional ranking |
| S+ | Meet tracking columns (2026-Times only) |

**Notes:**
- 100 IM is not a county or regional event — those columns are intentionally blank
- Ava's spreadsheet has no county/regional QT values yet — she is currently too young for the Middlesex county age groups
- The `2027-Times` tab shows what Margot needs to hit as a 13-year-old next season

---

## Claude Code shortcuts

| Command | What it does |
|---------|-------------|
| `/MargotSwimming` | Runs the main update after a gala — handles both girls |
| `/UpdateQualifyingTimes` | Annual QT update when new season PDFs are published |
