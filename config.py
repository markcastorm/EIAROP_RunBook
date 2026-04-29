# config.py
# EIAROP — EIA Russian Oil Production
# All constants, paths, column mappings, and settings

import os
import json
import pandas as pd
from datetime import datetime

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR  = os.path.join(BASE_DIR, 'downloads')
OUTPUT_DIR    = os.path.join(BASE_DIR, 'output')
MASTER_DIR    = os.path.join(BASE_DIR, 'Master data')
MASTER_FILE   = os.path.join(MASTER_DIR, 'Master_EIAROP_DATA.csv')
DATES_FILE    = os.path.join(BASE_DIR, 'release_dates.json')

# ── Timestamped folders ──────────────────────────────────────────────────────
RUN_TIMESTAMP = datetime.now().strftime('%Y%m%d_%H%M%S')

DOWNLOAD_RUN_DIR  = os.path.join(DOWNLOAD_DIR, RUN_TIMESTAMP)
OUTPUT_RUN_DIR    = os.path.join(OUTPUT_DIR, RUN_TIMESTAMP)
LATEST_OUTPUT_DIR = os.path.join(OUTPUT_DIR, 'latest')

# ── Source ────────────────────────────────────────────────────────────────────
BASE_URL = 'https://www.eia.gov/outlooks/steo/data.php?type=tables'

PROVIDER_NAME = 'EIA'
DATASET_NAME  = 'EIAROP'
DATA_UNIT     = 'barrels per day'

# ── Browser ───────────────────────────────────────────────────────────────────
HEADLESS_MODE      = True
WAIT_TIMEOUT       = 60
PAGE_LOAD_DELAY    = 5
DOWNLOAD_WAIT_TIME = 120

# ── Freshness check ─────────────────────────────────────────────────────────
# Set to False to bypass the "is data new?" check and always download
CHECK_FOR_NEW_DATA = True

# ── Output filenames ─────────────────────────────────────────────────────────
DATA_FILE_PATTERN = 'EIAROP_MONTHLY_DATA_{timestamp}.xls'
META_FILE_PATTERN = 'EIAROP_MONTHLY_META_{timestamp}.xls'
ZIP_FILE_PATTERN  = 'EIAROP_MONTHLY_{timestamp}.zip'

# ── Download settings ────────────────────────────────────────────────────────
MAX_DOWNLOAD_RETRIES = 3
RETRY_DELAY          = 3.0

# ── Downloaded file settings ─────────────────────────────────────────────────
# The Excel file downloaded from EIA STEO "All Tables"
STEO_FILENAME = 'STEO_m.xlsx'

# Tab to navigate to inside the downloaded workbook
TARGET_TAB = '3btab'

# Row label to find (code in col A, sublabel in col B)
TARGET_ROW_CODE     = 'papr_RS'
TARGET_ROW_SUBLABEL = 'Russia'

# ── Month mapping ────────────────────────────────────────────────────────────
MONTH_ABBR_TO_NUM = {
    'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04',
    'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08',
    'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12',
}

# =============================================================================
# SERIES DEFINITIONS (absolute — exact order matters)
# =============================================================================
# Single series for this runbook
# Format: (code_suffix, description)

SERIES_DEFINITIONS = [
    ('', 'EIA Russian Oil Production'),
]

# Derived lookup structures
SERIES_CODES = ['RUS.EIAROP.M']
SERIES_CODE_MNEMONICS = ['RUS.EIAROP.M']
SERIES_DESCRIPTIONS = [label for _, label in SERIES_DEFINITIONS]

# =============================================================================
# COLUMN MAPPING — Absolute column order for DATA output
# =============================================================================
# Row 0 (header 1): ["", "RUS.EIAROP.M"]
# Row 1 (header 2): ["", "EIA Russian Oil Production"]
# Data rows:         ["2022-01", 11.27760000000]

DATA_HEADER_ROW1 = [''] + SERIES_CODES
DATA_HEADER_ROW2 = [''] + SERIES_DESCRIPTIONS

# Values are written with full source precision (no truncation)

# =============================================================================
# META FILE CONFIGURATION
# =============================================================================

METADATA_COLUMNS = [
    'CODE',
    'CODE_MNEMONIC',
    'DESCRIPTION',
    'FREQUENCY',
    'MULTIPLIER',
    'AGGREGATION_TYPE',
    'UNIT_TYPE',
    'DATA_TYPE',
    'DATA_UNIT',
    'SEASONALLY_ADJUSTED',
    'ANNUALIZED',
    'STATE',
    'PROVIDER_MEASURE_URL',
    'PROVIDER',
    'SOURCE',
    'SOURCE_DESCRIPTION',
    'COUNTRY',
    'DATASET',
]

METADATA_DEFAULTS = {
    'FREQUENCY':            'M',
    'MULTIPLIER':           6.0,
    'AGGREGATION_TYPE':     'UNDEFINED',
    'UNIT_TYPE':            'FLOW',
    'DATA_TYPE':            'UNITS',
    'DATA_UNIT':            DATA_UNIT,
    'SEASONALLY_ADJUSTED':  'NSA',
    'ANNUALIZED':           False,
    'STATE':                'ACTIVE',
    'PROVIDER_MEASURE_URL': BASE_URL,
    'PROVIDER':             'EIA',
    'SOURCE':               'EIA',
    'SOURCE_DESCRIPTION':   'The U.S. Energy Information Administration',
    'COUNTRY':              'RUS',
    'DATASET':              DATASET_NAME,
}

# =============================================================================
# RELEASE DATES — persistence helpers
# =============================================================================

def load_release_dates():
    """Load previously saved release dates from JSON."""
    if os.path.exists(DATES_FILE):
        with open(DATES_FILE, 'r') as f:
            return json.load(f)
    return {}


def save_release_dates(dates_dict):
    """Save release dates to JSON."""
    with open(DATES_FILE, 'w') as f:
        json.dump(dates_dict, f, indent=2)


# =============================================================================
# MASTER DATA
# =============================================================================

def get_last_master_period():
    """Read the master CSV and return the last period (YYYY-MM) present."""
    if not os.path.exists(MASTER_FILE):
        return None
    df = pd.read_csv(MASTER_FILE)
    period_col = df.columns[0]
    periods = df[period_col].dropna().tolist()
    # Filter out non-date header rows
    date_periods = [p for p in periods if isinstance(p, str) and len(p) == 7 and '-' in p]
    return max(date_periods) if date_periods else None
