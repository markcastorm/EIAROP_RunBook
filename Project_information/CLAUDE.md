# EIAROP Runbook — Claude Code Context

This file gives a new Claude session full context to work on the EIAROP runbook without re-reading every file.

## Project Identity

- **Job Name**: EIAROP
- **Full Name**: EIA Russian Oil Production
- **Source**: U.S. Energy Information Administration (EIA) — Short-Term Energy Outlook (STEO)
- **URL**: https://www.eia.gov/outlooks/steo/data.php?type=tables
- **Frequency**: Monthly
- **Data Unit**: barrels per day (million barrels per day in source)
- **Country**: RUS (Russia)
- **Series Code**: `RUS.EIAROP.M`

## Architecture

Follows the standard SIMBA pipeline structure:

```
EIAROP_RunBook/
├── config.py             ← All constants, column mappings, paths, settings
├── scraper.py            ← Selenium stealth browser automation + XLSX download
├── extractor.py          ← openpyxl workbook parsing + dynamic header/row detection
├── file_generator.py     ← DATA/META/ZIP output + master CSV update
├── orchestrator.py       ← Wires download → extract → generate (returns 0/1)
├── main.py               ← Entry point: python main.py
├── release_dates.json    ← Persisted release dates for freshness checking
├── Master data/
│   └── Master_EIAROP_DATA.csv   ← Cumulative data (all periods)
├── downloads/
│   └── <timestamp>/             ← STEO_m.xlsx per run
├── output/
│   ├── <timestamp>/             ← DATA, META, ZIP per run
│   └── latest/                  ← Always holds the most recent output
└── Project_information/         ← Reference files, screenshots, this doc
```

## Pipeline Flow (orchestrator.py)

```
Step 1: scraper.download()
  → Navigate to EIA STEO tables page (Selenium stealth)
  → Scrape release dates from .pub_title div (Release Date, Forecast Completed, Next Release Date)
  → Save dates to release_dates.json
  → Compare against previous dates to detect new data (configurable: CHECK_FOR_NEW_DATA)
  → Dynamically find "All Tables" link by text content (not hardcoded URL)
  → Click link to trigger STEO_m.xlsx download (~1 MB)
  → Poll download dir for .xlsx (no .crdownload partials)
  → Returns xlsx_path

Step 2: extractor.STEOExtractor.extract(xlsx_path)
  → Open workbook with openpyxl (data_only=True)
  → Dynamically find target tab "3btab" (case-insensitive fallback)
  → Dynamically scan first 20 rows to find:
    - Year header row: row with 2+ integer values in range 1900-2100
    - Month header row: row with 6+ month abbreviations (Jan-Dec)
  → Build column-to-period map: {col_index: 'YYYY-MM'}
  → Dynamically scan all rows for TARGET_ROW_CODE="papr_RS" in column A
  → Extract all numeric values from that row, mapped to periods
  → Returns list of (period, value) tuples: [('2022-01', 11.2776), ...]

Step 3: file_generator.FileGenerator.generate_files(period_data, output_dir)
  → Creates DATA .xls (2 header rows + period data rows)
  → Creates META .xls (1 row, 18 metadata columns)
  → Creates ZIP (DATA + META)
  → Copies all to output/latest/
  → Appends new periods to Master_EIAROP_DATA.csv (deduplicates)
```

## config.py — Key Details

### Paths
- `BASE_DIR` = script directory (uses `os.path.dirname(os.path.abspath(__file__))`)
- `DOWNLOAD_DIR` / `OUTPUT_DIR` / `MASTER_DIR` — all under BASE_DIR
- `DATES_FILE` = `release_dates.json` — stores last-seen release dates
- All paths use `os.path.join()` — no hardcoded separators

### Timestamped Folders
- `RUN_TIMESTAMP` = `YYYYMMDD_HHMMSS` — computed once at import time
- `DOWNLOAD_RUN_DIR` = `downloads/<timestamp>/`
- `OUTPUT_RUN_DIR` = `output/<timestamp>/`
- `LATEST_OUTPUT_DIR` = `output/latest/`

### Browser Settings
- `HEADLESS_MODE = True` — always True in Docker (no display)
- `WAIT_TIMEOUT = 60` seconds
- `DOWNLOAD_WAIT_TIME = 120` seconds

### Freshness Check
- `CHECK_FOR_NEW_DATA = True` — compares release_date in JSON vs scraped page
- Set to `False` to bypass and always download/process

### Source File Settings
- `TARGET_TAB = '3btab'` — tab name to look for in the STEO workbook
- `TARGET_ROW_CODE = 'papr_RS'` — row code in column A
- `TARGET_ROW_SUBLABEL = 'Russia'` — expected sublabel in column B

### Single Series Definition
```
Code:        RUS.EIAROP.M
Mnemonic:    RUS.EIAROP.M
Description: EIA Russian Oil Production
```

### DATA Output Layout
```
Row 0: ["", "RUS.EIAROP.M"]              ← series code header
Row 1: ["", "EIA Russian Oil Production"] ← description header
Row 2: ["2022-01", "11.2776"]            ← data (full source precision)
Row 3: ["2022-02", "11.3308"]
...
```

### META Output Layout (18 columns)
```
CODE            = RUS.EIAROP.M
CODE_MNEMONIC   = RUS.EIAROP.M
DESCRIPTION     = EIA Russian Oil Production
FREQUENCY       = M
MULTIPLIER      = 6.0
AGGREGATION_TYPE= UNDEFINED
UNIT_TYPE       = FLOW
DATA_TYPE       = UNITS
DATA_UNIT       = barrels per day
SEASONALLY_ADJUSTED = NSA
ANNUALIZED      = False
STATE           = ACTIVE
PROVIDER_MEASURE_URL = https://www.eia.gov/outlooks/steo/data.php?type=tables
PROVIDER        = EIA
SOURCE          = EIA
SOURCE_DESCRIPTION = The U.S. Energy Information Administration
COUNTRY         = RUS
DATASET         = EIAROP
```

### Month Mapping
- `MONTH_ABBR_TO_NUM`: Jan→01, Feb→02, ..., Dec→12
- Used by extractor to convert month abbreviation columns to `YYYY-MM` periods

## scraper.py — Key Details

### Website Navigation Flow
1. **Load page**: `https://www.eia.gov/outlooks/steo/data.php?type=tables`
2. **Scrape dates**: Find `.pub_title` div, parse "Release Date:", "Forecast Completed:", "Next Release Date:" from pipe-delimited text
3. **Freshness check**: Compare scraped `release_date` against `release_dates.json`
4. **Find download link**: Search `a.ico_xls, a.xls` elements for text "All Tables"; fallback searches all `<a>` tags
5. **Trigger download**: JavaScript click on the link element; fallback navigates directly to URL
6. **Wait for file**: Poll download dir for `*.xlsx` with no `*.crdownload` partials

### Browser Setup
- Selenium with `selenium_stealth` library (graceful fallback if not installed)
- `--headless=new`, `--no-sandbox`, `--disable-dev-shm-usage`
- Suppresses `navigator.webdriver` flag via CDP
- Chrome version detection via `winreg` (Windows) or CLI (Linux)
- Human-like delays: `_human_delay(lo, hi)` with random sleep

### Download Link Detection (fully dynamic)
- Primary: `a.ico_xls, a.xls` elements containing text "All Tables"
- Fallback: any `<a>` tag with "All Tables" in text AND "STEO" in href
- Never hardcodes the download URL path

### Release Dates JSON Format
```json
{
  "release_date": "April 7, 2026",
  "forecast_completed": "April 6, 2026",
  "next_release_date": "May 12, 2026"
}
```

## extractor.py — Key Details

### Dynamic Header Discovery (no hardcoded row numbers)
The extractor scans the first 20 rows of the target worksheet to find:
1. **Year header row**: First row with 2+ integer values in range 1900-2100
   - Currently resolves to row 3: `{3: 2022, 15: 2023, 27: 2024, 39: 2025, 51: 2026, 63: 2027}`
2. **Month header row**: First row with 6+ month abbreviations (Jan-Dec)
   - Currently resolves to row 4: 72 month labels

### Column-to-Period Mapping
- Each month column is assigned to the nearest preceding year column
- Year columns appear at the start of each 12-month group
- Result: `{col_3: '2022-01', col_4: '2022-02', ..., col_74: '2027-12'}`

### Target Row Discovery
- Scans ALL rows in the worksheet (not just a range)
- Matches column A value against `config.TARGET_ROW_CODE` ("papr_RS")
- Logs the sublabel from column B for verification ("Russia")
- Currently found at row 26

### STEO Workbook Structure (STEO_m.xlsx)
- 28 tabs: Dates, Contents, 1tab, 2tab, 3atab, **3btab**, 3ctab, ...
- Tab **3btab**: "Non-OPEC Petroleum and Other Liquid Fuels Production (million barrels per day)"
- Row 1: Table title + Contents link
- Row 2: EIA attribution line
- Row 3: Year headers (2022, 2023, ..., 2027) at columns 3, 15, 27, 39, 51, 63
- Row 4: Month headers (Jan-Dec) repeating under each year — columns 3-74
- Row 5+: Data rows with code in col A, label in col B, values in cols 3-74
- Row 26: `papr_RS` | `Russia` | 72 data values

### Data Precision
- Source values range from 3 to 9 decimal places (e.g., `10.453` to `10.629521809`)
- Values are preserved exactly as Python floats — no truncation or padding
- `_format_value()` uses `Decimal(repr(float))` to guarantee lossless string conversion
- This applies to both DATA .xls output and master CSV

### Extracted Data Range
- 72 periods: 2022-01 through 2027-12
- First value: `('2022-01', 11.2776)`
- Last value: `('2027-12', 10.727541267)`
- Mix of historical actuals (fewer decimals) and forecast values (more decimals)

## file_generator.py — Key Details

### DATA File Layout (.xls via xlwt)
```
Row 0: ["", "RUS.EIAROP.M"]              ← series code
Row 1: ["", "EIA Russian Oil Production"] ← description
Row 2: ["2022-01", "11.2776"]            ← period + value (full precision string)
Row 3: ["2022-02", "11.3308"]
...
Row 73: ["2027-12", "10.727541267"]
```
- Values written as strings via `_format_value()` to preserve all decimal places
- 74 total rows (2 headers + 72 data periods)

### META File Layout (.xls via xlwt)
- Header row: 18 columns (CODE through DATASET)
- 1 data row with values from `config.METADATA_DEFAULTS`

### Master CSV Layout
```csv
,RUS.EIAROP.M
,EIA Russian Oil Production
2022-01,11.2776
2022-02,11.3308
...
```
- Description row is row index 0 (after header)
- New periods are appended; existing periods are skipped (deduplication)

### Output File Naming
- Timestamped: `EIAROP_MONTHLY_DATA_20260429_113411.xls`
- Latest: `EIAROP_MONTHLY_DATA_latest.xls`
- ZIP bundles DATA + META
- All three file types (DATA, META, ZIP) exist in both timestamped and latest folders

## Environment

- **Dev**: Windows 11, Python 3.11
- **Production**: Docker (Linux Ubuntu), Python 3.11
- All packages pre-installed in Docker image (no requirements.txt)
- Entry point: `python main.py`
- Uses `os.path.join()` everywhere — no hardcoded path separators
- `winreg` import is guarded with `sys.platform == 'win32'` check

### Required Packages (all pre-installed)
- `selenium`, `selenium-stealth` — browser automation
- `openpyxl` — reading STEO_m.xlsx
- `xlwt` — writing .xls output files
- `xlrd` — reading .xls (for verification only)
- `pandas` — master CSV handling
- Standard library: `os`, `sys`, `time`, `glob`, `json`, `logging`, `subprocess`, `random`, `shutil`, `zipfile`, `decimal`

## Known Behaviors and Edge Cases

1. **First run**: No `release_dates.json` exists — scraper treats data as new and downloads. This is correct behavior.
2. **Subsequent runs with same data**: If `CHECK_FOR_NEW_DATA=True` and the release date hasn't changed, the pipeline raises `RuntimeError('No new data available')` and exits with code 1. This is intentional.
3. **Bypass freshness check**: Set `CHECK_FOR_NEW_DATA=False` in config.py to always download and process regardless of release date.
4. **Duplicate periods in master**: The `update_master()` method checks existing periods and skips duplicates — safe to re-run.
5. **Year range changes**: If EIA adds more years (e.g., 2028), the extractor will pick them up automatically — the column map is built dynamically from whatever years appear in the header row.
6. **Tab structure changes**: If EIA renames the tab or the row code, update `TARGET_TAB` and `TARGET_ROW_CODE` in config.py.
7. **Decimal precision**: Values are never truncated. `_format_value()` uses `Decimal(repr(float))` so `10.629521809` stays exactly `"10.629521809"`, not `"10.6295218090"` or `"10.62952181"`.

## Reference Files (Project_information/)

| File | Purpose |
|------|---------|
| `information.txt` | Original requirements document |
| `EIAROP_DATA_20260408 - Sheet1.csv` | Manual/expected DATA output (reference for validation) |
| `EIAROP_DATA_20260408.xlsx` | XLSX version of manual data |
| `EIAROP_META_20250709.xlsx` | Reference META file (18 columns, 1 data row) |
| `EIAROP_RunBook.xlsx` | Project specification workbook |
| `STEO_m.xlsx` | Sample downloaded STEO workbook (for offline testing) |
| `image1.png` | Screenshot of EIA STEO tables page |
| `image2.png` | Screenshot of STEO_m.xlsx 3btab sheet layout |
| `sample.txt` | Raw HTML source of the EIA STEO page |
| `CLAUDE.md` | This file |

## Sibling Runbooks (same architecture)

- `D:\Projects\SIMBA-RUNBOOKS\TLIAD_Runbook\` — Taiwan Life Insurance (Annual, PDF source)
- `D:\Projects\SIMBA-RUNBOOKS\Runbook_RELPRCLVLINDX\` — Pipeline template with skeletons
