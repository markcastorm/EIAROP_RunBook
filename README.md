# EIAROP — EIA Russian Oil Production

SIMBA Automation Pipeline for scraping and processing Russian oil production data from the U.S. Energy Information Administration (EIA) Short-Term Energy Outlook (STEO).

## What It Does

1. Visits the [EIA STEO tables page](https://www.eia.gov/outlooks/steo/data.php?type=tables) using Selenium stealth
2. Scrapes release dates to detect new data (skips processing if unchanged)
3. Downloads the "All Tables" workbook (`STEO_m.xlsx`, ~1 MB)
4. Dynamically finds the `3btab` sheet and locates the `papr_RS` (Russia) row
5. Extracts all available monthly periods (currently 2022-01 through 2027-12)
6. Generates DATA (.xls), META (.xls), and ZIP output files
7. Updates the master CSV with new periods

## Quick Start

```bash
python main.py
```

Output appears in:
- `output/<timestamp>/` — timestamped run output
- `output/latest/` — always the most recent files
- `Master data/Master_EIAROP_DATA.csv` — cumulative data

## File Structure

```
EIAROP_RunBook/
├── config.py             ← Constants, paths, column mappings, metadata
├── scraper.py            ← Selenium stealth browser automation + download
├── extractor.py          ← Dynamic STEO workbook parsing
├── file_generator.py     ← DATA/META/ZIP generation + master CSV
├── orchestrator.py       ← Pipeline wiring (download → extract → generate)
├── main.py               ← Entry point
├── release_dates.json    ← Last-seen release dates (auto-generated)
├── Master data/          ← Cumulative master CSV
├── downloads/            ← Raw STEO_m.xlsx per run (timestamped)
├── output/               ← DATA/META/ZIP per run + latest
└── Project_information/  ← Reference files, screenshots, CLAUDE.md
```

## Configuration

All settings are in `config.py`. Key toggles:

| Setting | Default | Description |
|---------|---------|-------------|
| `HEADLESS_MODE` | `True` | Must be True in Docker (no display) |
| `CHECK_FOR_NEW_DATA` | `True` | Set `False` to bypass freshness check |
| `TARGET_TAB` | `'3btab'` | Tab name in STEO workbook |
| `TARGET_ROW_CODE` | `'papr_RS'` | Row label to extract |
| `WAIT_TIMEOUT` | `60` | Selenium wait timeout (seconds) |
| `DOWNLOAD_WAIT_TIME` | `120` | Max wait for file download (seconds) |

## Output Format

### DATA File (.xls)

| Row | Col 0 | Col 1 |
|-----|-------|-------|
| 0 | *(empty)* | `RUS.EIAROP.M` |
| 1 | *(empty)* | `EIA Russian Oil Production` |
| 2 | `2022-01` | `11.2776` |
| 3 | `2022-02` | `11.3308` |
| ... | ... | ... |

Values preserve full source precision (3-9 decimal places, no truncation).

### META File (.xls)

Single data row with 18 columns: CODE, CODE_MNEMONIC, DESCRIPTION, FREQUENCY (M), MULTIPLIER (6.0), AGGREGATION_TYPE (UNDEFINED), UNIT_TYPE (FLOW), DATA_TYPE (UNITS), DATA_UNIT (barrels per day), SEASONALLY_ADJUSTED (NSA), ANNUALIZED (False), STATE (ACTIVE), PROVIDER_MEASURE_URL, PROVIDER (EIA), SOURCE (EIA), SOURCE_DESCRIPTION, COUNTRY (RUS), DATASET (EIAROP).

## How Dynamic Detection Works

The pipeline avoids hardcoded row/column positions:

- **Tab**: Found by name match (case-insensitive fallback)
- **Year headers**: First row in top 20 with 2+ integers in range 1900-2100
- **Month headers**: First row in top 20 with 6+ month abbreviations
- **Data row**: Full worksheet scan for `papr_RS` in column A
- **Download link**: Text search for "All Tables" across page links
- **Release dates**: Text parsing from `.pub_title` div content

If EIA adds years, shifts headers, or moves the data row, the extractor adapts automatically.

## Environment

- **Dev**: Windows 11, Python 3.11
- **Production**: Docker (Linux Ubuntu), Python 3.11
- All packages pre-installed in Docker image
- Entry point: `python main.py`
- Docker mounts: script folder → `/app`, output → `/app/output`

## Dependencies (pre-installed)

selenium, selenium-stealth, openpyxl, xlwt, xlrd, pandas, plus standard library modules.

## Claude Context

For starting a new Claude Code session on this project, see `Project_information/CLAUDE.md` — it contains full technical context for every file, data structures, edge cases, and reference material paths.
