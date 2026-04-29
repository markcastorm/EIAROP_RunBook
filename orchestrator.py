# orchestrator.py
# Wires the full EIAROP pipeline: download → extract → generate.

import sys
import logging

import config
from scraper import download
from extractor import STEOExtractor
from file_generator import FileGenerator

logger = logging.getLogger(__name__)


def main():
    """Run the full pipeline. Returns 0 on success, 1 on failure."""
    logging.basicConfig(
        stream=sys.stdout,
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )

    # Silence noisy third-party loggers
    for noisy in ('selenium', 'selenium.webdriver', 'urllib3',
                   'urllib3.connectionpool', 'openpyxl'):
        logging.getLogger(noisy).setLevel(logging.WARNING)

    try:
        logger.info('=== EIAROP pipeline started ===')
        logger.info(f'Timestamp: {config.RUN_TIMESTAMP}')
        logger.info(f'Source:    {config.BASE_URL}')
        logger.info(f'Master:   {config.MASTER_FILE}')

        # ── Step 1: Download ─────────────────────────────────────────────
        logger.info('Step 1: Downloading STEO workbook from EIA...')
        xlsx_path = download()
        logger.info(f'Downloaded: {xlsx_path}')

        # ── Step 2: Extract ──────────────────────────────────────────────
        logger.info('Step 2: Extracting data from STEO workbook...')
        extractor = STEOExtractor()
        period_data = extractor.extract(xlsx_path)

        if not period_data:
            logger.error('No data extracted — aborting')
            return 1

        logger.info(
            f'Extracted: {len(period_data)} periods '
            f'({period_data[0][0]} to {period_data[-1][0]})'
        )

        # ── Step 3: Generate output files ────────────────────────────────
        logger.info('Step 3: Generating output files...')
        generator = FileGenerator()
        output_files = generator.generate_files(
            period_data, config.OUTPUT_RUN_DIR
        )

        # ── Summary ──────────────────────────────────────────────────────
        logger.info('=== EIAROP pipeline completed successfully ===')
        logger.info(f'Output dir:  {config.OUTPUT_RUN_DIR}')
        logger.info(f'Latest dir:  {config.LATEST_OUTPUT_DIR}')
        logger.info(f'DATA: {output_files["data_file"]}')
        logger.info(f'META: {output_files["meta_file"]}')
        logger.info(f'ZIP:  {output_files["zip_file"]}')

        return 0

    except Exception as e:
        logger.exception(f'Pipeline failed: {e}')
        return 1
