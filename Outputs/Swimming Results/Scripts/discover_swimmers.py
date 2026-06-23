"""
Automated swimmer discovery for swimmingresults.org.

Uses Playwright (headless Chromium) to search the event rankings page filtered
by club, age group, sex, course, and stroke. Iterates through events to build
a comprehensive list of all registered swimmers at each club/YOB group.

Usage:
    python3 discover_swimmers.py               # check for missing/new swimmers
    python3 discover_swimmers.py --update      # also add new swimmers to config JSONs

Clubs and age groups checked:
    Brompton SC              — Female 2013, 2014, 2015
    Chelsea & Westminster SC — Female 2016, 2017, 2018
"""

import json, re, time, sys
from pathlib import Path
from playwright.sync_api import sync_playwright

SCRIPTS_DIR = Path(__file__).parent
ROOT_DIR    = SCRIPTS_DIR.parent
MARGOT_DIR  = ROOT_DIR / "Margot"
AVA_DIR     = ROOT_DIR / "Ava"

UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"

CURRENT_YEAR = 2026

# Club codes on swimmingresults.org
CLUB_CODES = {
    "Brompton":  "BRBLMDXL",
    "Chelsea":   "CWSLMDXL",
}

GROUPS = [
    {"config": MARGOT_DIR / "margot_swimmers_2013.json", "club": "Brompton", "yob": 2013},
    {"config": MARGOT_DIR / "margot_swimmers.json",      "club": "Brompton", "yob": 2014},
    {"config": MARGOT_DIR / "margot_swimmers_2015.json", "club": "Brompton", "yob": 2015},
    {"config": AVA_DIR    / "ava_swimmers_2016.json",    "club": "Chelsea",  "yob": 2016},
    {"config": AVA_DIR    / "ava_swimmers.json",         "club": "Chelsea",  "yob": 2017},
    {"config": AVA_DIR    / "ava_swimmers_2018.json",    "club": "Chelsea",  "yob": 2018},
]

# All events to check (stroke code, course)
EVENTS = [
    ("1",  "S"), ("1",  "L"),   # 50m Free SC / LC
    ("2",  "S"), ("2",  "L"),   # 100m Free
    ("3",  "S"), ("3",  "L"),   # 200m Free
    ("7",  "S"), ("7",  "L"),   # 50m Breast
    ("8",  "S"), ("8",  "L"),   # 100m Breast
    ("13", "S"), ("13", "L"),   # 50m Back
    ("14", "S"), ("14", "L"),   # 100m Back
    ("10", "S"), ("10", "L"),   # 50m Fly
    ("11", "S"), ("11", "L"),   # 100m Fly
    ("16", "S"), ("16", "L"),   # 200m IM
]

STROKE_LABELS = {
    "1": "50Free", "2": "100Free", "3": "200Free",
    "7": "50Breast", "8": "100Breast",
    "13": "50Back", "14": "100Back",
    "10": "50Fly", "11": "100Fly",
    "16": "200IM",
}


def yob_to_age_group(yob, search_year=None):
    """Return the AgeGroup select value for a swimmer born in yob in the given search year."""
    year = search_year if search_year is not None else CURRENT_YEAR
    age = year - yob
    return f"{age:02d}"


def load_config(path):
    if path.exists():
        return json.loads(path.read_text())
    return {"swimmers": []}


def save_config(path, config):
    path.write_text(json.dumps(config, indent=2))


def find_all_club_swimmers(page, club_key, yob):
    """
    Query the event rankings page across multiple events to find all swimmers
    registered at the club with the given YOB. Returns {tiref: name} dict.

    Searches the current year and previous year, using the age correct for that
    year (target_year - yob) so that results from different years don't bleed
    across birth-year groups.
    """
    club_code    = CLUB_CODES[club_key]
    found        = {}  # tiref -> name
    target_years = [str(CURRENT_YEAR), str(CURRENT_YEAR - 1)]

    for stroke, course in EVENTS:
        for target_year in target_years:
            age_group = yob_to_age_group(yob, search_year=int(target_year))
            try:
                page.goto("https://www.swimmingresults.org/eventrankings/", timeout=30000)
                page.wait_for_load_state("networkidle", timeout=15000)

                page.select_option('select[name="Pool"]',       course)
                page.select_option('select[name="Stroke"]',     stroke)
                page.select_option('select[name="Sex"]',        "F")
                page.select_option('select[name="TargetYear"]', target_year)
                page.select_option('select[name="AgeGroup"]',   age_group)
                page.select_option('select[name="AgeAt"]',      "D")  # 31st December
                page.click('input[type="radio"][value="O"]')           # Club level
                page.select_option('select[name="TargetClub"]', club_code)
                page.fill('input[name="RecordsToView"]',        "200")

                page.click('input[type="submit"]')
                page.wait_for_load_state("networkidle", timeout=20000)

                links = page.query_selector_all('a[href*="tiref"]')
                for link in links:
                    href = link.get_attribute("href") or ""
                    m    = re.search(r'tiref=(\d+)', href)
                    if not m:
                        continue
                    tiref = m.group(1)
                    if tiref not in found:
                        name = link.inner_text().strip()
                        if name:
                            found[tiref] = name

                time.sleep(0.3)

            except Exception as e:
                label = f"{STROKE_LABELS.get(stroke, stroke)} {course} {target_year}"
                print(f"      {label}: error — {e}")
                continue

    return found


def main():
    do_update   = "--update"   in sys.argv
    only_bsc    = "--brompton" in sys.argv
    only_cwsc   = "--cwsc"     in sys.argv

    print("Swimmer Discovery — swimmingresults.org")
    print("=" * 50)

    with sync_playwright() as p:
        browser  = p.chromium.launch(headless=True)
        context  = browser.new_context(user_agent=UA, locale="en-GB")
        page     = context.new_page()

        any_new  = False

        for group in GROUPS:
            if only_bsc  and group["club"] != "Brompton": continue
            if only_cwsc and group["club"] != "Chelsea":  continue
            config   = load_config(group["config"])
            existing = {s["name"]: s["member_id"] for s in config.get("swimmers", [])}
            yob      = group["yob"]
            club_key = group["club"]
            label    = f"{club_key} Female {yob} (YOB {yob}, age {CURRENT_YEAR - yob})"

            print(f"\nScanning {label} — {len(existing)} in config...")

            discovered = find_all_club_swimmers(page, club_key, yob)
            print(f"  Site total: {len(discovered)} swimmers found")

            new_swimmers = []
            for tiref, name in discovered.items():
                if name not in existing:
                    new_swimmers.append({"name": name, "member_id": tiref})
                    print(f"  + NEW: {name} (tiref={tiref})")

            departed = []
            for name in existing:
                if name not in discovered.values():
                    departed.append(name)
            if departed:
                print(f"  ? Not found on site ({len(departed)}): {', '.join(departed)}")
                print(f"    (may have no times in the last 2 years, or left the club)")

            if new_swimmers:
                any_new = True
                if do_update:
                    for sw in new_swimmers:
                        config["swimmers"].append(sw)
                    save_config(group["config"], config)
                    print(f"  → Config updated (+{len(new_swimmers)} swimmers)")
            else:
                print(f"  ✓ No new swimmers")

        browser.close()

    print("\n" + "=" * 50)
    if any_new and not do_update:
        print("New swimmers found. Run with --update to add them to config files.")
    elif not any_new:
        print("All config files are up to date.")
    else:
        print("Config files updated. Run swimming_results.py --refresh to pull their times.")


if __name__ == "__main__":
    main()
