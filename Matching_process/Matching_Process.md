# Matching Process Description: WIPO patent data merged with startup data

Date of WIPO searches: 19 October 2026. Source: WIPO PATENTSCOPE (patentscope.wipo.int), signed-in account.

## 1. Goal
Link each patent record WIPO holds for a startup to the correct row of the Top-20 startup sheet, and count the records per startup.

## 2. Step-by-step

**Step 1. Search by applicant name.** For each of the 20 startups, one or more searches were run in the Applicant field using the form `PA:("Company Name")`, with all offices and no date limit. Searches were run on the registered or brand name; where a short or generic name returned unrelated companies, a narrower registered-style name was used (for example `River mobility`, `Matter Motor Works`, `Smartron India`). Every query, date, result count and outcome is in `WIPO_Search_Log.xlsx`.

**Step 2. Screen before downloading.** The Applicant column of the first results was read for each search. A search was downloaded only if the applicants were the startup. Searches returning only other companies were not downloaded and are logged as "0 relevant results" (see section 4).

**Step 3. Download raw data.** For the 10 startups with relevant results, "Download > 10,000 results" produced one Excel file each, with the fields Application Number and Date, Publication Number and Date, Country, Title, IPC and Applicants. The files are kept unedited in `raw_wipo/`. Each file carries its own query and timestamp in its header lines.

**Step 4. Merge.** `wipo_merge.py` reads all 10 files (732 rows) into one table and records the source file and query for each row.

**Step 5. Normalise and match applicant names.** The Applicants cell was split on ";" into separate parties, because some filings list co-applicants or inventors. Each party name was normalised (lower case, punctuation removed, legal suffixes such as Pvt, Private, Ltd, Limited, Pte, B.V. removed) and compared with the startup's registered-name alias (`Alias_Table`). An exact normalised match was accepted. A near match scoring 0.90 or above was accepted automatically, 0.80 to 0.90 was flagged for manual review, and below 0.80 was rejected. In this data every accepted row was an exact normalised match, so nothing needed manual review on name similarity.

**Step 6. Cross-check.** The matched startup for every row was compared with the startup the file was searched for. There were no disagreements.

**Step 7. Join to the startup sheet.** Counts per startup are formulas over the merged table, placed beside the funding, category and rank columns of the ranked startup sheet (`Startup_Summary`).

## 3. Results
- 732 rows downloaded; **730 matched**, **2 rejected** (both CN filings by "Nantong Ola Electric Appliance Co., Ltd.", a Chinese appliance maker, found in the Ola Electric file; name similarity 0.59).
- The merge was validated by checking that Rows = Matched + Rejected in every file and that row counts equal the WIPO result counts in the search log (`File_Checks`). A first run wrongly rejected one genuine Ather row written as "PVT.LTD." with no space; the name normaliser was corrected and the merge rerun.
- Matched joint filings (21 rows) list other parties, such as Indian Oil Corporation on 3 Chakr filings and a named inventor on some Carbon Clean filings. They are kept and flagged in the "Other parties on filing" column.
- Name variants merged under one startup: Ather Energy Pvt. Ltd. and Limited; Ola Electric Mobility Private Limited, Limited, Ltd and B.V.; Carbon Clean Solutions Limited, Ltd and Pvt. Ltd.; SUN Mobility Pte Ltd.

## 4. Startups with no patent data in the merged table
Searches were run, but no relevant results were found, so no file exists: ACME Solar, Euler Motors, Battery Smart, Simple Energy, SolarSquare, Bounce, Agratas, EKA Mobility, EV Co, ChargeZone. Several searches returned unrelated companies with similar names, which were rejected on the Applicant column, for example Bounce (176 hits), EV Co (240 hits, e.g. a Korean circuit maker), Battery Smart (hits from Chinese and German companies), Simple Energy (UK and US companies), ACME (US companies).

## 5. Limitations
1. **Name-based matching only.** A patent filed under a parent, founder or other group entity would be missed, and a zero means "none found under the names searched", not "no patents exist".
2. **Records, not inventions.** The same invention filed in several countries appears as several records. Counts are WIPO patent records. Carbon Clean has 104 records but 100 unique publication numbers.
3. **Filed vs granted cannot be separated.** The export has no kind or grant code, and PATENTSCOPE mainly lists published applications. All records are treated as filed and published.
4. **Publication lag.** Applications are typically published about 18 months after filing, so recent startups (Agratas, EKA) may have unpublished filings.
5. **Group entities.** SUN Mobility's records are filed under a Singapore entity and Carbon Clean's mainly under a UK entity, so most are non-Indian office filings. They were matched on name and subject area.
6. **Subject area.** Smartron's records (2016 to 2020) are mainly phone and IoT topics, not environmental technology.
