# Incubator Website Collector (Wayback Machine)

Collects the **homepage** and **About-us page** of each incubator's website:

| Version | What is saved |
|---|---|
| **Current** | The live pages, fetched on the day the script was run (19 October 2026) |
| **Archived** | The **earliest available capture in each calendar month** on the Internet Archive's Wayback Machine, going back **September 2026 → January 2022** (57 months) |

Everything is saved as raw HTML, and a `manifest.csv` records the outcome for every incubator × page × month.

---

## 1. Files in this package

```
Q1_Submission/
├── Q1_Methodology_and_Ranking_Justification.pdf   <- search methodology + ranking justification
├── Q1_Incubator_Research.xlsx                     <- incubator details, ranking, all 32 candidates
├── README.md                                      <- this file
├── code/
│   ├── wayback_scraper.py                         <- the annotated Python 3 script
│   ├── requirements.txt                           <- Python dependencies
│   └── incubators.csv                             <- the final Top 20 (name, homepage_url, about_url)
├── output/                                        <- the scraped HTML, manifest.csv, run.log
├── n1/                                            <- later run: JSS ATEB STEP
├── n2/                                            <- later run: Amity Innovation Incubator
└── n3/                                            <- later run: Agri-Business Incubator (ABI-ICRISAT)
```

**The scraped HTML is included** in `output/`. It was collected by running the script on a machine with internet access
on 18-19 October 2026.

`incubators.csv` lists the 20 incubators of the final ranking in `Q1_Incubator_Research.xlsx` (sheet *Top 20 Ranking*).

---

## 2. Setup

Requires Python 3.9+.

```bash
cd code
pip install -r requirements.txt
```

## 3. Running it

Run from the `code/` folder:

```bash
# Test on one incubator first (recommended)
python wayback_scraper.py --input incubators.csv --out ../output --contact you@university.edu --only "T-Hub"

# Full run (all 20 incubators)
python wayback_scraper.py --input incubators.csv --out ../output --contact you@university.edu
```

`--contact` is required. It goes into the User-Agent header so the Internet Archive can identify (and, if needed, email) the person running the bot.

**Expected runtime.** The script asks the archive's index for *all months at once* (one query per URL, not one per month), so the cost is dominated by downloading pages: roughly 2 pages × 57 months × 20 incubators ≈ up to ~2,300 downloads. At the default 1 s delay that is about 40-60 minutes, depending on how many months have captures and on retries.

**Note:** `manifest.csv` is rewritten at the start of every run. If you run the script again with `--only`, point `--out` at a different folder (or back up `manifest.csv` first), then merge.

### Options

| Option | Default | Meaning |
|---|---|---|
| `--input` | `incubators.csv` | CSV **or** XLSX with columns `name`, `homepage_url`, optional `about_url` |
| `--out` | `output` | Output folder |
| `--contact` | *(required)* | Your email, for the User-Agent |
| `--start` / `--end` | `2022-01` / `2026-09` | Month range, `YYYY-MM` |
| `--delay` | `1.0` | Seconds between requests. Don't go much lower |
| `--only` | – | Run only incubators whose name contains this text |
| `--skip-current` | off | Skip the live-site fetch (archive only) |
| `--relaxed-status` | off | Also accept redirect (301/302/…) and "revisit" captures (see 5.1) |

### Input format

```
name,homepage_url,about_url
T-Hub,https://www.t-hub.co/,
Maker Village,http://makervillage.in/,https://makervillage.in/about
```

`about_url` is optional. If you know the exact About URL, filling it in is the most reliable option; this is especially recommended for incubators that are a *section* of a larger institution's site (e.g. IIT Madras Incubation Cell, Kerala Startup Mission), where automatic discovery may land on the institution's generic About page.

---

## 4. Output layout

```
output/
├── manifest.csv                       <- one row per incubator x page x month
├── run.log                            <- detailed log
└── t_hub/
    ├── homepage/
    │   ├── current_20261019.html
    │   ├── 2026-09_20260911104355.html     <- {month}_{wayback timestamp}.html
    │   ├── 2026-08_20260806155917.html
    │   └── ...
    └── about/
        └── (same naming)
```

The timestamp in the filename (`YYYYMMDDhhmmss`) is the exact moment of the capture, so you can see how early in the month it was. Paths in `manifest.csv` are relative to the root of the submission folder.

### manifest.csv columns

`incubator, page_type (homepage|about), month (YYYY-MM or "current"), status, snapshot_timestamp, requested_url, archived_url, file, bytes, sha256, note`

| status | Meaning |
|---|---|
| `ok` | Downloaded and saved |
| `skipped_exists` | File already on disk from an earlier run (not re-downloaded) |
| `no_snapshot` | The archive has no (acceptable) capture of that page in that month |
| `error` | A capture exists but could not be downloaded, or the index query failed (see `note`) |
| `missing` | (current about page only) no About link was found on the live homepage |

The `sha256` column lets you spot months where the page did not actually change.

---

## 5. How it works

1. **Read the input list** (CSV or XLSX).
2. **Current pages.** Fetch the live homepage; use `about_url` if given, otherwise search the homepage for a same-site link whose text or path looks like "About / About us / Who we are / Our story" (shortest path wins). Fetch that page too.
3. **Archived homepage.** One query to the Wayback CDX index returns the earliest capture of the homepage in *each* month of the range (`collapse=timestamp:6` keeps the first row per `YYYYMM`). Each is downloaded in its original, unmodified form (`…/web/<timestamp>id_/<url>`: no Wayback toolbar, no rewritten links).
4. **Archived about page.** URLs are tried in this order until every month is covered:
   1. the `about_url` from your input,
   2. the About URL found on the live homepage,
   3. common guesses (`/about-us`, `/about`, `/about.html`, `/who-we-are`, …), relative to the site root and to the homepage path,
   4. for any months still uncovered: the About link found inside *that month's archived homepage* (handles sites that were redesigned and renamed their pages).
5. **Resume.** Months already saved on disk are skipped, so you can stop (Ctrl-C) and re-run at any time. The manifest is rebuilt each run and includes the skipped files.
6. **Politeness.** Delay before every request, descriptive User-Agent, exponential back-off on HTTP 429/5xx and network errors.

### 5.1 "Earliest available" and status codes

By default only captures where the server returned **HTTP 200** count. Some sites are mostly archived as redirects (e.g. `http→https`), which can produce months marked `no_snapshot` even though the site was archived. If you see many unexpected gaps, re-run with `--relaxed-status` (delete nothing; existing files are kept and only missing months are filled).

---

## 6. Results of the run

The script was run on 18-19 October 2026. The live pages were fetched on 19 October 2026.

- **Collection in stages.** The first full run used an earlier candidate list. After the eligibility rule was finalised, three incubators that entered the final Top 20 were collected in separate, later runs, each into its own output folder:
  - `n1/` – JSS ATEB STEP
  - `n2/` – Amity Innovation Incubator
  - `n3/` – Agri-Business Incubator (ABI-ICRISAT)

  These folders were left in place as the raw output of those runs (each has its own `manifest.csv` and `run.log`). Their HTML was also copied into `output/`, so `output/` holds all 20 final incubators. The main `output/manifest.csv` covers the first run and JSS ATEB STEP; the Amity and ICRISAT rows are in `n2/manifest.csv` and `n3/manifest.csv`.
- **Extra folders.** `output/` also contains three organisations (100X.VC, Peak XV Surge, Atal Community Innovation Centres) that were scraped before the eligibility rule excluded them from the Top 20 (investment funds and umbrella schemes). They are kept for completeness and are not part of the ranking.
- **Where to look.** `manifest.csv` gives the status of every incubator × page × month, so per-incubator coverage can be counted directly from it.

## 7. Known limitations (worth stating in your methods section)

- **Missing months are normal.** The Wayback Machine does not capture every site every month; `no_snapshot` means "not archived", not "site was down". Coverage varies widely between incubators: for example Marwari Catalysts has very few archived homepage months, and Amrita TBI, JioGenNext and Zone Startups India have no archived About pages.
- **JSS ATEB STEP has no archived pages.** Only its live homepage was saved. No About link was found on the live site and the archive holds no captures of either page for January 2022 to September 2026.
- **Live fetches that failed.** The live homepage could not be fetched for NSRCEL, IIM Calcutta Innovation Park, Maker Village and Marwari Catalysts (the sites closed the connection or timed out). Their archived pages were collected normally.
- **No About link on the live homepage.** For several incubators no About link could be found automatically on the live homepage (recorded as `missing` in the manifest). Where an `about_url` was supplied in `incubators.csv`, it was used for the archived pages.
- **Failed archived downloads.** A few captures listed in the Wayback index could not be downloaded (recorded as `error`, e.g. WE Hub 2024-01 and GUSEC 2024-06 and 2025-03).
- **Guessed About pages can be wrong.** A site may return a normal-looking page (HTTP 200) for a guessed URL that isn't really an About page. Spot-check, and set `about_url` for any incubator where it matters.
- **JavaScript-rendered sites.** Archived HTML is the raw source. Content that a site loads with JavaScript may be absent from the file.
- **Embedded assets** (images, CSS) are not downloaded; only the HTML.
- **"Earliest in the month"** is the earliest capture the archive indexed, which may be any day of that month.
- **Live pages** are saved once per run as `current_<date>.html`.
- Respect each website's terms and the Internet Archive's usage policies; keep the delay at ≥ 1 s.

## 8. Troubleshooting

| Symptom | Fix |
|---|---|
| Many `error` rows with `CDX index query failed` | Archive is rate-limiting or down; wait and re-run (it resumes) or raise `--delay` to 2-3 |
| `No usable rows in …` | Check column headers are exactly `name`, `homepage_url` |
| Everything `no_snapshot` for one incubator | Check the URL is correct (try `https://web.archive.org/web/*/<url>` in a browser); try `--relaxed-status` |
| Live page `error` | Site blocks bots or is down; archived data is unaffected |
| `FileNotFoundError: incubators.csv` | Run the command from inside the `code/` folder |