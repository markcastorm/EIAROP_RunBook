# file_generator.py
# Generate DATA (.xls), META (.xls), and ZIP output files for EIAROP.
# Also updates the master CSV with new data.

import os
import shutil
import zipfile
import logging
from decimal import Decimal, ROUND_HALF_UP

import xlwt
import pandas as pd

import config

logger = logging.getLogger(__name__)


def _format_value(value):
    """
    Format a float value preserving ALL source decimal places with no
    truncation.  Uses Decimal(repr(float)) so that every significant
    digit present in the Python float survives the round-trip into the
    output file.  Returns a string.
    """
    d = Decimal(repr(value))
    # Normalize removes trailing zeros; we keep them by returning the
    # full Decimal string which already has exact precision.
    return str(d)


class FileGenerator:
    """
    Generates SIMBA-standard output files for EIAROP:
      - DATA file: 2 header rows (code + description) + data rows (period, value)
      - META file: one row per series with standard metadata columns
      - ZIP file: contains both DATA and META
    Output goes to timestamped folder + 'latest' folder.
    Master CSV is updated with new periods.
    """

    def __init__(self):
        self.logger = logger

    # ─────────────────────────────────────────────────────────────────────
    # DATA file
    # ─────────────────────────────────────────────────────────────────────

    def create_data_file(self, period_data, output_path):
        """
        Create the DATA Excel file.

        Layout:
            Row 0: ["", "RUS.EIAROP.M"]
            Row 1: ["", "EIA Russian Oil Production"]
            Row 2: ["2022-01", 11.27760000000]
            Row 3: ["2022-02", 11.33080000000]
            ...

        Args:
            period_data: list of (period, value) tuples sorted by period
            output_path: full path for the output .xls file
        """
        self.logger.info('Creating DATA file...')

        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('DATA')

        # Row 0: series codes
        for col_idx, header in enumerate(config.DATA_HEADER_ROW1):
            sheet.write(0, col_idx, header)

        # Row 1: series descriptions
        for col_idx, header in enumerate(config.DATA_HEADER_ROW2):
            sheet.write(1, col_idx, header)

        # Data rows — preserve full source precision as strings to avoid
        # float display truncation in Excel; Decimal ensures no digits are lost
        for row_offset, (period, value) in enumerate(period_data):
            row_idx = row_offset + 2
            sheet.write(row_idx, 0, period)
            # Convert via Decimal to capture every digit from the source
            sheet.write(row_idx, 1, _format_value(value))

        workbook.save(output_path)

        self.logger.info(
            f'DATA file saved: {output_path}  |  '
            f'{len(period_data)} periods'
        )
        return output_path

    # ─────────────────────────────────────────────────────────────────────
    # META file
    # ─────────────────────────────────────────────────────────────────────

    def create_meta_file(self, output_path):
        """Create the META Excel file. One row per series."""
        self.logger.info('Creating META file...')

        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('META')

        # Header row
        for col_idx, col_name in enumerate(config.METADATA_COLUMNS):
            sheet.write(0, col_idx, col_name)

        # Data row (single series for EIAROP)
        for series_idx, (suffix, label) in enumerate(config.SERIES_DEFINITIONS):
            row_idx = series_idx + 1

            row_data = {
                'CODE':                 config.SERIES_CODES[series_idx],
                'CODE_MNEMONIC':        config.SERIES_CODE_MNEMONICS[series_idx],
                'DESCRIPTION':          label,
                'FREQUENCY':            config.METADATA_DEFAULTS['FREQUENCY'],
                'MULTIPLIER':           config.METADATA_DEFAULTS['MULTIPLIER'],
                'AGGREGATION_TYPE':     config.METADATA_DEFAULTS['AGGREGATION_TYPE'],
                'UNIT_TYPE':            config.METADATA_DEFAULTS['UNIT_TYPE'],
                'DATA_TYPE':            config.METADATA_DEFAULTS['DATA_TYPE'],
                'DATA_UNIT':            config.METADATA_DEFAULTS['DATA_UNIT'],
                'SEASONALLY_ADJUSTED':  config.METADATA_DEFAULTS['SEASONALLY_ADJUSTED'],
                'ANNUALIZED':           config.METADATA_DEFAULTS['ANNUALIZED'],
                'STATE':                config.METADATA_DEFAULTS['STATE'],
                'PROVIDER_MEASURE_URL': config.METADATA_DEFAULTS['PROVIDER_MEASURE_URL'],
                'PROVIDER':             config.METADATA_DEFAULTS['PROVIDER'],
                'SOURCE':               config.METADATA_DEFAULTS['SOURCE'],
                'SOURCE_DESCRIPTION':   config.METADATA_DEFAULTS['SOURCE_DESCRIPTION'],
                'COUNTRY':              config.METADATA_DEFAULTS['COUNTRY'],
                'DATASET':              config.METADATA_DEFAULTS['DATASET'],
            }

            for col_idx, col_name in enumerate(config.METADATA_COLUMNS):
                value = row_data.get(col_name, '')
                sheet.write(row_idx, col_idx, value)

        workbook.save(output_path)

        self.logger.info(
            f'META file saved: {output_path}  |  '
            f'{len(config.SERIES_DEFINITIONS)} series'
        )
        return output_path

    # ─────────────────────────────────────────────────────────────────────
    # ZIP file
    # ─────────────────────────────────────────────────────────────────────

    def create_zip_file(self, data_file, meta_file, zip_path):
        """Bundle DATA and META files into a single ZIP."""
        self.logger.info(f'Creating ZIP: {zip_path}')

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.write(data_file, os.path.basename(data_file))
            zf.write(meta_file, os.path.basename(meta_file))

        self.logger.info('ZIP created')
        return zip_path

    # ─────────────────────────────────────────────────────────────────────
    # Master data update
    # ─────────────────────────────────────────────────────────────────────

    def update_master(self, period_data):
        """
        Append new periods to the master CSV.

        The master CSV layout:
            Row 0 (header 1): ["", "RUS.EIAROP.M"]
            Row 1 (header 2): ["", "EIA Russian Oil Production"]
            Row 2+: ["YYYY-MM", value_string]

        Args:
            period_data: list of (period, value) tuples
        """
        master_path = config.MASTER_FILE
        os.makedirs(os.path.dirname(master_path), exist_ok=True)
        self.logger.info(f'Updating master: {master_path}')

        if os.path.exists(master_path):
            # Read existing master — skip the description header row
            df = pd.read_csv(master_path)
            period_col = df.columns[0]
            existing_periods = set(df[period_col].dropna().astype(str).tolist())
        else:
            # Create new master with proper headers
            cols = [''] + config.SERIES_CODES
            df = pd.DataFrame(columns=cols)
            # Add description row
            desc_row = {cols[0]: '', cols[1]: config.SERIES_DESCRIPTIONS[0]}
            df = pd.concat([df, pd.DataFrame([desc_row])], ignore_index=True)
            period_col = df.columns[0]
            existing_periods = set()

        # Add new periods that don't already exist
        new_rows = []
        for period, value in period_data:
            if period in existing_periods:
                continue
            new_rows.append({
                df.columns[0]: period,
                df.columns[1]: _format_value(value),
            })

        if new_rows:
            df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
            df.to_csv(master_path, index=False)
            self.logger.info(f'Master updated: {len(new_rows)} new periods added')
        else:
            self.logger.info('Master unchanged — all periods already present')

    # ─────────────────────────────────────────────────────────────────────
    # Generate all outputs
    # ─────────────────────────────────────────────────────────────────────

    def generate_files(self, period_data, output_dir):
        """
        Generate DATA, META, and ZIP files.

        Args:
            period_data: list of (period, value) tuples sorted by period
            output_dir: timestamped output directory

        Returns:
            dict with paths to all created files
        """
        os.makedirs(output_dir, exist_ok=True)

        timestamp = config.RUN_TIMESTAMP

        data_filename = config.DATA_FILE_PATTERN.format(timestamp=timestamp)
        meta_filename = config.META_FILE_PATTERN.format(timestamp=timestamp)
        zip_filename  = config.ZIP_FILE_PATTERN.format(timestamp=timestamp)

        data_path = os.path.join(output_dir, data_filename)
        meta_path = os.path.join(output_dir, meta_filename)
        zip_path  = os.path.join(output_dir, zip_filename)

        # Create files
        self.create_data_file(period_data, data_path)
        self.create_meta_file(meta_path)
        self.create_zip_file(data_path, meta_path, zip_path)

        # Copy to 'latest' folder
        latest_dir = config.LATEST_OUTPUT_DIR
        os.makedirs(latest_dir, exist_ok=True)

        latest_data = os.path.join(latest_dir, 'EIAROP_MONTHLY_DATA_latest.xls')
        latest_meta = os.path.join(latest_dir, 'EIAROP_MONTHLY_META_latest.xls')
        latest_zip  = os.path.join(latest_dir, 'EIAROP_MONTHLY_latest.zip')

        shutil.copy2(data_path, latest_data)
        shutil.copy2(meta_path, latest_meta)
        shutil.copy2(zip_path, latest_zip)

        self.logger.info(f'Files copied to latest: {latest_dir}')

        # Update master CSV
        self.update_master(period_data)

        return {
            'data_file':   data_path,
            'meta_file':   meta_path,
            'zip_file':    zip_path,
            'latest_data': latest_data,
            'latest_meta': latest_meta,
            'latest_zip':  latest_zip,
        }
