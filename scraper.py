# scraper.py
# Downloads the STEO "All Tables" Excel file from the EIA Short-Term Energy
# Outlook page using Selenium stealth, and scrapes release dates.

import os
import sys
import time
import glob
import logging
import subprocess
import random

import config

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# Chrome version detection (Windows dev + Linux Docker)
# ─────────────────────────────────────────────────────────────────────────────

def get_chrome_version():
    """Detect Chrome major version — works on Windows (dev) and Linux (Docker)."""
    if sys.platform == 'win32':
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r'Software\Google\Chrome\BLBeacon',
            )
            return winreg.QueryValueEx(key, 'version')[0].split('.')[0]
        except Exception:
            pass
    for cmd in ['google-chrome', 'google-chrome-stable',
                'chromium', 'chromium-browser']:
        try:
            out = subprocess.check_output(
                [cmd, '--version'], stderr=subprocess.DEVNULL
            ).decode()
            return out.strip().split()[-1].split('.')[0]
        except Exception:
            continue
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Helper: wait utilities
# ─────────────────────────────────────────────────────────────────────────────

def _human_delay(lo=0.4, hi=1.2):
    """Small random pause to mimic human speed."""
    time.sleep(random.uniform(lo, hi))


def _wait_and_click(driver, by, value, timeout=None, description='element'):
    """Wait for an element to be clickable, scroll into view, then click."""
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    timeout = timeout or config.WAIT_TIMEOUT
    wait = WebDriverWait(driver, timeout)
    el = wait.until(EC.element_to_be_clickable((by, value)))
    driver.execute_script('arguments[0].scrollIntoView({block:"center"});', el)
    _human_delay()
    el.click()
    logger.debug(f'Clicked: {description}')
    return el


def _wait_for(driver, by, value, timeout=None, description='element'):
    """Wait for element presence and return it."""
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    timeout = timeout or config.WAIT_TIMEOUT
    wait = WebDriverWait(driver, timeout)
    el = wait.until(EC.presence_of_element_located((by, value)))
    logger.debug(f'Found: {description}')
    return el


# ─────────────────────────────────────────────────────────────────────────────
# Build driver
# ─────────────────────────────────────────────────────────────────────────────

def _build_driver(download_dir):
    """Create a Selenium stealth Chrome driver configured for file download."""
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    try:
        from selenium_stealth import stealth
    except ImportError:
        stealth = None

    abs_dl = os.path.abspath(download_dir)
    os.makedirs(abs_dl, exist_ok=True)

    opts = Options()
    if config.HEADLESS_MODE:
        opts.add_argument('--headless=new')
        opts.add_argument('--disable-gpu')

    opts.add_argument('--no-sandbox')
    opts.add_argument('--disable-dev-shm-usage')
    opts.add_argument('--window-size=1920,1080')
    opts.add_argument('--disable-blink-features=AutomationControlled')
    opts.add_argument(
        'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/131.0.0.0 Safari/537.36'
    )

    prefs = {
        'download.default_directory': abs_dl,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': False,
    }
    opts.add_experimental_option('prefs', prefs)
    opts.add_experimental_option('excludeSwitches', ['enable-automation'])
    opts.add_experimental_option('useAutomationExtension', False)

    driver = webdriver.Chrome(options=opts)
    driver.set_page_load_timeout(config.WAIT_TIMEOUT * 2)

    # Apply selenium-stealth if available
    if stealth is not None:
        stealth(
            driver,
            languages=['en-US', 'en'],
            vendor='Google Inc.',
            platform='Win32',
            webgl_vendor='Intel Inc.',
            renderer='Intel Iris OpenGL Engine',
            fix_hairline=True,
        )
        logger.info('Selenium stealth applied')

    # Suppress navigator.webdriver flag
    driver.execute_cdp_cmd(
        'Page.addScriptToEvaluateOnNewDocument',
        {'source': 'Object.defineProperty(navigator,"webdriver",'
                    '{get:()=>undefined})'},
    )

    logger.info(f'Chrome driver ready — download dir: {abs_dl}')
    return driver


# ─────────────────────────────────────────────────────────────────────────────
# Scrape release dates from page
# ─────────────────────────────────────────────────────────────────────────────

def _scrape_release_dates(driver):
    """
    Scrape Release Date, Forecast Completed, and Next Release Date
    from the pub_title section of the STEO page.

    Returns dict: {
        'release_date': '...',
        'forecast_completed': '...',
        'next_release_date': '...',
    }
    """
    from selenium.webdriver.common.by import By

    logger.info('Scraping release dates...')

    dates = {}

    try:
        pub_div = driver.find_element(By.CSS_SELECTOR, '.pub_title')
        text = pub_div.text
        logger.debug(f'pub_title text: {text}')

        # Parse dates from the text content
        # Format: "Release Date: April 7, 2026 | Forecast Completed: April 6, 2026 | Next Release Date: May 12, 2026"
        for line in text.replace('\n', ' ').split('|'):
            line = line.strip()
            if 'Release Date:' in line and 'Next' not in line:
                dates['release_date'] = line.split('Release Date:')[1].strip()
            elif 'Forecast Completed:' in line:
                dates['forecast_completed'] = line.split('Forecast Completed:')[1].strip()
            elif 'Next Release Date:' in line:
                dates['next_release_date'] = line.split('Next Release Date:')[1].strip()
    except Exception as e:
        logger.warning(f'Could not scrape release dates: {e}')

    logger.info(f'Release dates scraped: {dates}')
    return dates


def _is_data_new(current_dates):
    """
    Compare current release dates against previously saved dates.
    Returns True if data is new (or no previous dates saved).
    """
    previous_dates = config.load_release_dates()

    if not previous_dates:
        logger.info('No previous release dates found — treating as new data')
        return True

    prev_release = previous_dates.get('release_date', '')
    curr_release = current_dates.get('release_date', '')

    if curr_release and curr_release != prev_release:
        logger.info(
            f'New data detected: previous release={prev_release}, '
            f'current release={curr_release}'
        )
        return True

    logger.info(f'Data unchanged (release_date={curr_release})')
    return False


# ─────────────────────────────────────────────────────────────────────────────
# Find and click the "All Tables" download link
# ─────────────────────────────────────────────────────────────────────────────

def _find_all_tables_link(driver):
    """
    Dynamically find the 'All Tables' download link on the STEO page.
    Returns the href URL of the link.
    """
    from selenium.webdriver.common.by import By

    logger.info('Looking for "All Tables" download link...')

    # Look for links with class ico_xls or xls that contain "All Tables" text
    links = driver.find_elements(By.CSS_SELECTOR, 'a.ico_xls, a.xls')
    for link in links:
        text = (link.text or '').strip()
        if 'All Tables' in text:
            href = link.get_attribute('href')
            logger.info(f'Found "All Tables" link: {href}')
            return href

    # Fallback: search all links for "All Tables"
    all_links = driver.find_elements(By.TAG_NAME, 'a')
    for link in all_links:
        text = (link.text or '').strip()
        href = link.get_attribute('href') or ''
        if 'All Tables' in text and 'STEO' in href:
            logger.info(f'Found "All Tables" link (fallback): {href}')
            return href

    raise RuntimeError('"All Tables" download link not found on page')


def _click_download_link(driver, download_url):
    """
    Click the All Tables download link to trigger the file download.
    Uses JavaScript click as a reliable approach for download links.
    """
    from selenium.webdriver.common.by import By

    logger.info(f'Triggering download: {download_url}')

    # Find the link element by href and click it
    links = driver.find_elements(By.CSS_SELECTOR, 'a')
    for link in links:
        href = link.get_attribute('href') or ''
        if href == download_url:
            driver.execute_script('arguments[0].scrollIntoView({block:"center"});', link)
            _human_delay()
            driver.execute_script('arguments[0].click();', link)
            logger.info('Download link clicked')
            return

    # Fallback: navigate directly to the URL
    logger.info('Link element not found — navigating directly to download URL')
    driver.get(download_url)


# ─────────────────────────────────────────────────────────────────────────────
# Wait for downloaded file
# ─────────────────────────────────────────────────────────────────────────────

def _wait_for_downloaded_file(download_dir, extension='.xlsx', timeout=None):
    """
    Poll the download directory until a file with the given extension appears
    (and no .crdownload / .tmp partial files remain).
    """
    timeout = timeout or config.DOWNLOAD_WAIT_TIME
    abs_dir = os.path.abspath(download_dir)
    logger.info(f'Waiting for {extension} download in: {abs_dir} (timeout={timeout}s)')

    start = time.time()
    while time.time() - start < timeout:
        partials = (
            glob.glob(os.path.join(abs_dir, '*.crdownload'))
            + glob.glob(os.path.join(abs_dir, '*.tmp'))
        )
        target_files = glob.glob(os.path.join(abs_dir, f'*{extension}'))

        if target_files and not partials:
            target_files.sort(key=os.path.getmtime, reverse=True)
            result = target_files[0]
            file_size = os.path.getsize(result)
            logger.info(f'Download complete: {os.path.basename(result)} ({file_size:,} bytes)')
            return result

        time.sleep(2)

    raise RuntimeError(f'Download timed out after {timeout}s — no {extension} file in {abs_dir}')


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def download():
    """
    Full EIA STEO scraping workflow:
    1. Navigate to the STEO tables page
    2. Scrape release dates and check freshness
    3. Find and click the "All Tables" download link
    4. Wait for the STEO_m.xlsx file to download
    5. Return the path to the downloaded file

    Returns:
        str: path to the downloaded STEO_m.xlsx file

    Raises:
        RuntimeError: if download fails or data is not new
    """
    download_dir = config.DOWNLOAD_RUN_DIR
    os.makedirs(download_dir, exist_ok=True)

    driver = None
    try:
        driver = _build_driver(download_dir)

        # Navigate to the STEO tables page
        logger.info(f'Navigating to: {config.BASE_URL}')
        driver.get(config.BASE_URL)
        time.sleep(config.PAGE_LOAD_DELAY)
        _human_delay(1.0, 2.0)
        logger.info('Page loaded')

        # Scrape release dates
        current_dates = _scrape_release_dates(driver)

        # Check if data is new
        if config.CHECK_FOR_NEW_DATA:
            if not _is_data_new(current_dates):
                logger.info('No new data available — skipping download')
                raise RuntimeError('No new data available (release date unchanged)')
        else:
            logger.info('Freshness check bypassed (CHECK_FOR_NEW_DATA=False)')

        # Save current release dates
        config.save_release_dates(current_dates)

        # Find the "All Tables" download link
        download_url = _find_all_tables_link(driver)

        # Click to trigger download
        _click_download_link(driver, download_url)

        # Wait for the file to finish downloading
        xlsx_path = _wait_for_downloaded_file(download_dir, extension='.xlsx')

        logger.info(f'Downloaded STEO file: {xlsx_path}')
        return xlsx_path

    finally:
        if driver:
            driver.quit()
            logger.info('Browser closed')
