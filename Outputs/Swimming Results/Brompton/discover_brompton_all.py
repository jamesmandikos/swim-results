"""
discover_brompton_all.py — Find all Brompton SC swimmers across all age groups.

Uses Playwright to query event rankings, filtered by club, YOB, and sex.
Writes config files to ../Brompton/config/brompton_{f|m}_{yob}.json.

Usage:
    python3 discover_brompton_all.py           # preview only (no writes)
    python3 discover_brompton_all.py --update  # write/update config files
"""

import json, re, time, sys
from pathlib import Path
from playwright.sync_api import sync_playwright

SCRIPTS_DIR  = Path(__file__).parent
BROMPTON_DIR = SCRIPTS_DIR.parent / "Brompton"
CONFIG_DIR   = BROMPTON_DIR / "config"
CONFIG_DIR.mkdir(parents=True, exist_ok=True)

CLUB_CODE   = "BRBLMDXL"
CURRENT_YEAR = 2026

UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
      "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

# Key events to scan — enough to catch all active swimmers without excessive loads.
# (stroke_code, course_code) — SC=S, LC=L
SCAN_EVENTS = [
    ("1",  "S"),   # 50 Free SC
    ("2",  "S"),   # 100 Free SC
    ("3",  "S"),   # 200 Free SC
    ("4",  "S"),   # 400 Free SC
    ("13", "S"),   # 50 Back SC
    ("14", "S"),   # 100 Back SC
    ("7",  "S"),   # 50 Breast SC
    ("8",  "S"),   # 100 Breast SC
    ("10", "S"),   # 50 Fly SC
    ("11", "S"),   # 100 Fly SC
    ("16", "S"),   # 200 IM SC
    ("1",  "L"),   # 50 Free LC
    ("2",  "L"),   # 100 Free LC
    ("14", "L"),   # 100 Back LC
    ("8",  "L"),   # 100 Breast LC
]

STROKE_LABELS = {
    "1":"50 Free","2":"100 Free","3":"200 Free","4":"400 Free",
    "13":"50 Back","14":"100 Back","7":"50 Breast","8":"100 Breast",
    "10":"50 Fly","11":"100 Fly","16":"200 IM",
}
COURSE_LABELS = {"S":"SC","L":"LC"}

# YOBs and genders to scan
YOBS   = list(range(2009, 2018))   # 2009–2017
GENDERS = [("F","Female"), ("M","Male")]


def config_path(gender_code, yob):
    return CONFIG_DIR / f"brompton_{gender_code.lower()}_{yob}.json"


def load_config(path, gender, yob):
    if path.exists():
        return json.loads(path.read_text())
    return {"club": "Brompton SC", "gender": gender, "year_of_birth": yob, "swimmers": []}


def save_config(path, config):
    path.write_text(json.dumps(config, indent=2))


def find_swimmers(page, sex_code, yob):
    """Return {tiref: name} for all swimmers found at Brompton for this sex/YOB."""
    age_group = f"{CURRENT_YEAR - yob:02d}"
    found = {}

    for stroke, course in SCAN_EVENTS:
        label = f"{STROKE_LABELS[stroke]} {COURSE_LABELS[course]}"
        try:
            page.goto("https://www.swimmingresults.org/eventrankings/", timeout=30000)
            page.wait_for_load_state("networkidle", timeout=20000)

            page.select_option('select[name="Pool"]',       course)
            page.select_option('select[name="Stroke"]',     stroke)
            page.select_option('select[name="Sex"]',        sex_code)
            page.select_option('select[name="TargetYear"]', "2026")
            page.select_option('select[name="AgeGroup"]',   age_group)
            page.select_option('select[name="AgeAt"]',      "D")
            page.click('input[type="radio"][value="O"]')
            page.select_option('select[name="TargetClub"]', CLUB_CODE)
            page.fill('input[name="RecordsToView"]',        "200")
            page.click('input[type="submit"]')
            page.wait_for_load_state("networkidle", timeout=20000)

            links = page.query_selector_all('a[href*="tiref"]')
            before = len(found)
            for link in links:
                href  = link.get_attribute("href") or ""
                m     = re.search(r'tiref=(\d+)', href)
                if not m:
                    continue
                tiref = m.group(1)
                if tiref not in found:
                    name = link.inner_text().strip()
                    if name:
                        found[tiref] = name
            new = len(found) - before
            print(f"      {label}: {len(links)} entries, {new} new swimmers")
            time.sleep(0.5)

        except Exception as e:
            print(f"      {label}: ERROR — {e}")

    return found


def main():
    do_update = "--update" in sys.argv

    print("Brompton SC — Full Swimmer Discovery")
    print("=" * 50)
    print(f"Mode: {'UPDATE config files' if do_update else 'PREVIEW only'}")
    print()

    summary = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(user_agent=UA, locale="en-GB")
        page    = context.new_page()

        for yob in YOBS:
            for sex_code, gender_label in GENDERS:
                age   = CURRENT_YEAR - yob
                key   = f"{sex_code.lower()}_{yob}"
                path  = config_path(sex_code, yob)
                cfg   = load_config(path, gender_label, yob)
                existing = {s["name"]: s["member_id"] for s in cfg.get("swimmers", [])}

                print(f"\n── {gender_label} {yob} (age {age}) — {len(existing)} in config ──")

                discovered = find_swimmers(page, sex_code, yob)
                print(f"  Site: {len(discovered)} swimmers found")

                new_swimmers = []
                for tiref, name in discovered.items():
                    if name not in existing:
                        new_swimmers.append({"name": name, "member_id": tiref})
                        print(f"  + NEW: {name} (id={tiref})")

                departed = [n for n in existing if n not in discovered.values()]
                if departed:
                    print(f"  ? Missing from site ({len(departed)}): {', '.join(departed)}")

                total = len(existing) + len(new_swimmers)
                summary[key] = {"gender": gender_label, "yob": yob, "age": age,
                                 "existing": len(existing), "new": len(new_swimmers),
                                 "total": total}

                if new_swimmers and do_update:
                    for sw in new_swimmers:
                        cfg["swimmers"].append(sw)
                    save_config(path, cfg)
                    print(f"  → Saved {total} swimmers to {path.name}")
                elif not path.exists() and do_update:
                    save_config(path, cfg)
                    print(f"  → Created {path.name} (0 swimmers found yet)")

        browser.close()

    print("\n" + "=" * 50)
    print("SUMMARY")
    print(f"{'Group':<12} {'Gender':<8} {'Age':>4} {'Existing':>9} {'New':>5} {'Total':>6}")
    print("-" * 50)
    grand = 0
    for key, s in sorted(summary.items()):
        print(f"{key:<12} {s['gender']:<8} {s['age']:>4} {s['existing']:>9} {s['new']:>5} {s['total']:>6}")
        grand += s["total"]
    print(f"\nTotal swimmers across all groups: {grand}")

    if not do_update:
        print("\nRun with --update to write config files.")


if __name__ == "__main__":
    main()
