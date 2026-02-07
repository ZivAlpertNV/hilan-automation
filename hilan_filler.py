#!/usr/bin/env python3
"""
Hilan Attendance Hours Filler
=============================
Automates filling work hours in the Hilan (חילן) attendance system.
Supports NVIDIA's Hilan instance at https://nvidia.net.hilan.co.il

Usage:
    python hilan_filler.py --user <EMPLOYEE_NUM> --password <PASSWORD> [options]

Author: Auto-generated
"""

# Must be set BEFORE importing playwright so it finds browsers
# installed alongside the package (via `playwright install` with PLAYWRIGHT_BROWSERS_PATH=0)
import os
os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", "0")

import argparse
import calendar
import logging
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import holidays
from playwright.sync_api import (
    Browser,
    BrowserContext,
    Page,
    Playwright,
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
)

# ==============================================================================
# Configuration & Constants
# ==============================================================================

HILAN_BASE_URL = "https://nvidia.net.hilan.co.il"
HILAN_LOGIN_URL = f"{HILAN_BASE_URL}/login"
HILAN_HOME_URL = f"{HILAN_BASE_URL}/Hilannetv2/ng/management/home"
HILAN_ATTENDANCE_URL = f"{HILAN_BASE_URL}/Hilannetv2/Attendance/calendarpage.aspx"

# --- Login Page Selectors (verified from saved HTML) ---
SEL_USERNAME_INPUT = "#user_nm"
SEL_PASSWORD_INPUT = "#password_nm"
SEL_LOGIN_BUTTON = "button[type='submit']"

# --- Attendance Calendar Page Selectors (verified from saved HTML) ---
# The page has: (1) a mini-calendar on the left for selecting days,
#                (2) a reports grid on the right showing entry/exit fields.
# You MUST select days in the calendar first, then the grid shows their rows.

# Calendar: individual day cells use aria-label with day number
# e.g., td[aria-label="1"], td[aria-label="2"], ...
# Classes: cDIES = regular workday, cHD = holiday/weekend, cED = error/missing

# Reports grid: rows are <tr> with id containing "row_N"
# Date cell span has title like "01/02 Sun"
# ManualEntry input id pattern:  ..._cellOf_ManualEntry_EmployeeReports_row_N_0_ManualEntry_EmployeeReports_row_N_0
# ManualExit  input id pattern:  ..._cellOf_ManualExit_EmployeeReports_row_N_0_ManualExit_EmployeeReports_row_N_0
# Symbol dropdown id pattern:    ..._cellOf_Symbol.SymbolId_EmployeeReports_row_N_0_Symbol.SymbolId_EmployeeReports_row_N_0

# Save button
SEL_SAVE_BUTTON = "input[id$='btnSave'][value='Save']"

# Screenshot directory
SCREENSHOT_DIR = Path(__file__).parent / "screenshots"


# ==============================================================================
# Logging Setup
# ==============================================================================

def setup_logging(verbose: bool = False) -> logging.Logger:
    """Configure and return the logger."""
    log_level = logging.DEBUG if verbose else logging.INFO
    logger = logging.getLogger("hilan_filler")
    logger.setLevel(log_level)

    # Console handler with UTF-8 encoding for Hebrew support
    import io
    console_handler = logging.StreamHandler(
        io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    )
    console_handler.setLevel(log_level)
    formatter = logging.Formatter(
        "[%(asctime)s] %(levelname)-8s %(message)s",
        datefmt="%H:%M:%S",
    )
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    return logger


# ==============================================================================
# CLI Argument Parsing
# ==============================================================================

def parse_args() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Hilan Attendance Hours Filler - מילוי שעות אוטומטי בחילן",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Fill current month with default hours (09:00-18:00)
  python hilan_filler.py -u 12345 -p mypassword

  # Fill specific month with custom hours
  python hilan_filler.py -u 12345 -p mypassword --month 1 --year 2026 --start-time 08:30 --end-time 17:30

  # Dry run - see which days would be filled
  python hilan_filler.py -u 12345 -p mypassword --dry-run

  # Mark specific dates as present (e.g., 12th and 19th of the month)
  python hilan_filler.py -u 12345 -p mypassword --present-dates 12,19

  # Combine weekly + specific present days
  python hilan_filler.py -u 12345 -p mypassword --present-days 2 --present-dates 12,19

  # Run in headless mode (no browser window)
  python hilan_filler.py -u 12345 -p mypassword --headless
        """,
    )

    # Required arguments
    parser.add_argument(
        "-u", "--user",
        required=True,
        help="Employee number (מספר עובד)",
    )
    parser.add_argument(
        "-p", "--password",
        required=True,
        help="Password (סיסמה)",
    )

    # Optional arguments
    parser.add_argument(
        "--start-time",
        default="09:00",
        help="Entry time (שעת כניסה) - default: 09:00",
    )
    parser.add_argument(
        "--end-time",
        default="18:00",
        help="Exit time (שעת יציאה) - default: 18:00",
    )
    parser.add_argument(
        "--project",
        required=True,
        help="Project code/name to fill in the Project column (e.g., '12086 - AGUR IC')",
    )
    parser.add_argument(
        "--present-days",
        default=None,
        help=(
            "Days of the week you come to the office EVERY week (comma-separated). "
            "These days will be marked as 'presence' instead of 'w.home'. "
            "Format: 1=Sun, 2=Mon, 3=Tue, 4=Wed, 5=Thu. "
            "Example: --present-days 2,4 (Monday and Wednesday in office every week)"
        ),
    )
    parser.add_argument(
        "--present-dates",
        default=None,
        help=(
            "Specific days of the month you were in the office (comma-separated). "
            "These days will be marked as 'presence' instead of 'w.home'. "
            "Supports: individual days (1,2,3), ranges (10-14), or both (5,12,20-22). "
            "Example: --present-dates 5,12,19 (present on the 5th, 12th, and 19th)"
        ),
    )
    parser.add_argument(
        "--vacation",
        default=None,
        help=(
            "Days of the month to mark as vacation (no hours filled). "
            "Supports: individual days (1,2,3), ranges (1-6), or both (1-3,15,20-22). "
            "Example: --vacation 8-12 (vacation from 8th to 12th)"
        ),
    )
    parser.add_argument(
        "--sick-days",
        default=None,
        help=(
            "Days of the month to mark as sick. "
            "1-2 days: 'sick day declaration' (no file needed). "
            "3+ days: 'sick' (requires --sick-file). "
            "Supports: individual days (1,2,3), ranges (1-6), or both (1-3,15). "
            "Example: --sick-days 8-10"
        ),
    )
    parser.add_argument(
        "--sick-file",
        default=None,
        help="Path to sick certificate file (required when --sick-days has 3+ days)",
    )
    parser.add_argument(
        "--month",
        type=int,
        default=None,
        help="Month to fill (1-12) - default: current month",
    )
    parser.add_argument(
        "--year",
        type=int,
        default=None,
        help="Year to fill - default: current year",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        default=False,
        help="Run browser in headless mode (no visible window)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        default=False,
        help="Only print which days would be filled, don't actually do it",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        default=False,
        help="Enable verbose/debug logging",
    )
    parser.add_argument(
        "--slow-mo",
        type=int,
        default=500,
        help="Slow down Playwright actions by this many ms (default: 500)",
    )

    args = parser.parse_args()

    # Validate time format
    for time_arg, name in [(args.start_time, "--start-time"), (args.end_time, "--end-time")]:
        try:
            parts = time_arg.split(":")
            h, m = int(parts[0]), int(parts[1])
            if not (0 <= h <= 23 and 0 <= m <= 59):
                raise ValueError
        except (ValueError, IndexError):
            parser.error(f"{name} must be in HH:MM format (e.g., 09:00)")

    # Default month/year to current
    today = date.today()
    if args.month is None:
        args.month = today.month
    if args.year is None:
        args.year = today.year

    # Validate month/year
    if not (1 <= args.month <= 12):
        parser.error("--month must be between 1 and 12")
    if args.year < 2020 or args.year > 2030:
        parser.error("--year must be between 2020 and 2030")

    # Parse present-days into a set of Python weekday numbers
    # User format: 1=Sun, 2=Mon, 3=Tue, 4=Wed, 5=Thu
    # Python weekday(): 0=Mon, 1=Tue, 2=Wed, 3=Thu, 4=Fri, 5=Sat, 6=Sun
    user_to_python_weekday = {1: 6, 2: 0, 3: 1, 4: 2, 5: 3}  # 1=Sun→6, 2=Mon→0, etc.
    args.present_weekdays = set()
    if args.present_days:
        try:
            for d in args.present_days.split(","):
                d = d.strip()
                if not d:
                    continue
                day_num = int(d)
                if day_num not in user_to_python_weekday:
                    parser.error(f"--present-days: invalid day {day_num}. Use 1=Sun, 2=Mon, 3=Tue, 4=Wed, 5=Thu")
                args.present_weekdays.add(user_to_python_weekday[day_num])
        except ValueError:
            parser.error("--present-days must be comma-separated numbers (e.g., 2,4)")

    # Parse present-dates into a set of day-of-month numbers
    present_dates_raw = args.present_dates
    args.present_dates = set()
    if present_dates_raw:
        try:
            for part in present_dates_raw.split(","):
                part = part.strip()
                if not part:
                    continue
                if "-" in part:
                    start, end = part.split("-", 1)
                    for d in range(int(start.strip()), int(end.strip()) + 1):
                        args.present_dates.add(d)
                else:
                    args.present_dates.add(int(part))
        except ValueError:
            parser.error("--present-dates must be comma-separated days or ranges (e.g., 5,12,20-22)")

    # Parse sick days into a set of day-of-month numbers
    args.sick_days_set = set()
    if args.sick_days:
        try:
            for part in args.sick_days.split(","):
                part = part.strip()
                if not part:
                    continue
                if "-" in part:
                    start, end = part.split("-", 1)
                    for d in range(int(start.strip()), int(end.strip()) + 1):
                        args.sick_days_set.add(d)
                else:
                    args.sick_days_set.add(int(part))
        except ValueError:
            parser.error("--sick-days must be comma-separated days or ranges (e.g., 1-3,15)")

    # Validate sick-file requirement
    if len(args.sick_days_set) > 2 and not args.sick_file:
        parser.error("--sick-file is required when --sick-days has more than 2 days")
    if args.sick_file:
        sick_path = Path(args.sick_file)
        if not sick_path.exists():
            parser.error(f"--sick-file: file not found: {args.sick_file}")

    # Parse vacation days into a set of day-of-month numbers
    args.vacation_days = set()
    if args.vacation:
        try:
            for part in args.vacation.split(","):
                part = part.strip()
                if not part:
                    continue
                if "-" in part:
                    # Range: "1-6" → {1, 2, 3, 4, 5, 6}
                    start, end = part.split("-", 1)
                    for d in range(int(start.strip()), int(end.strip()) + 1):
                        args.vacation_days.add(d)
                else:
                    args.vacation_days.add(int(part))
        except ValueError:
            parser.error("--vacation must be comma-separated days or ranges (e.g., 1-6,15,20-22)")

    # Validate no overlap between vacation and sick days
    if args.vacation_days and args.sick_days_set:
        overlap = args.vacation_days & args.sick_days_set
        if overlap:
            parser.error(
                f"--vacation and --sick-days overlap on days: {sorted(overlap)}. "
                f"A day cannot be both vacation and sick."
            )

    # Validate no overlap between present-dates and vacation/sick days
    if args.present_dates and args.vacation_days:
        overlap = args.present_dates & args.vacation_days
        if overlap:
            parser.error(
                f"--present-dates and --vacation overlap on days: {sorted(overlap)}. "
                f"A day cannot be both present and vacation."
            )
    if args.present_dates and args.sick_days_set:
        overlap = args.present_dates & args.sick_days_set
        if overlap:
            parser.error(
                f"--present-dates and --sick-days overlap on days: {sorted(overlap)}. "
                f"A day cannot be both present and sick."
            )

    return args


# ==============================================================================
# Workday Calculation
# ==============================================================================

def get_israeli_holidays(year: int) -> set:
    """Get all Israeli holidays for the given year."""
    il_holidays = holidays.Israel(years=year)
    return set(il_holidays.keys())


def get_workdays(year: int, month: int) -> list[date]:
    """
    Calculate all workdays (Sunday-Thursday) in the given month,
    excluding Israeli holidays.
    Includes ALL days of the month (including future) so the script
    can fix incorrect symbols on future days too.
    """
    holiday_dates = get_israeli_holidays(year)

    # Get the number of days in the month
    _, num_days = calendar.monthrange(year, month)

    today = date.today()
    workdays = []

    for day_num in range(1, num_days + 1):
        d = date(year, month, day_num)

        # Sunday=6, Monday=0, Tuesday=1, Wednesday=2, Thursday=3, Friday=4, Saturday=5
        # In Python: Monday=0, Sunday=6
        # We want Sunday(6) through Thursday(3)
        # Actually in Python's weekday(): Monday=0, Tuesday=1, ..., Sunday=6
        # We want: Sunday(6), Monday(0), Tuesday(1), Wednesday(2), Thursday(3)
        weekday = d.weekday()
        is_workday = weekday in (6, 0, 1, 2, 3)  # Sun, Mon, Tue, Wed, Thu

        if not is_workday:
            continue

        # Skip holidays
        if d in holiday_dates:
            continue

        workdays.append(d)

    return workdays


def print_workdays_summary(workdays: list[date], year: int, month: int,
                           start_time: str, end_time: str, logger: logging.Logger):
    """Print a summary of workdays that will be filled."""
    month_name = calendar.month_name[month]
    day_names = {0: "Mon", 1: "Tue", 2: "Wed", 3: "Thu", 4: "Fri", 5: "Sat", 6: "Sun"}
    day_names_he = {0: "שני", 1: "שלישי", 2: "רביעי", 3: "חמישי", 4: "שישי", 5: "שבת", 6: "ראשון"}

    holiday_dates = get_israeli_holidays(year)

    logger.info("=" * 60)
    logger.info(f"  Month: {month_name} {year} ({month}/{year})")
    logger.info(f"  Hours: {start_time} - {end_time}")
    logger.info(f"  Workdays to fill: {len(workdays)}")
    logger.info("=" * 60)

    for d in workdays:
        day_name = day_names[d.weekday()]
        day_name_he = day_names_he[d.weekday()]
        logger.info(f"  {d.strftime('%d/%m/%Y')}  {day_name} ({day_name_he})  {start_time}-{end_time}")

    # Show skipped days
    _, num_days = calendar.monthrange(year, month)
    today = date.today()
    skipped_holidays = []
    skipped_weekends = []
    skipped_future = []

    for day_num in range(1, num_days + 1):
        d = date(year, month, day_num)
        if d in [wd for wd in workdays]:
            continue
        weekday = d.weekday()
        if d > today:
            skipped_future.append(d)
        elif d in holiday_dates:
            holiday_name = holidays.Israel(years=year).get(d, "Holiday")
            skipped_holidays.append((d, holiday_name))
        elif weekday in (4, 5):  # Fri, Sat
            skipped_weekends.append(d)

    if skipped_holidays:
        logger.info("-" * 60)
        logger.info("  Skipped holidays:")
        for d, name in skipped_holidays:
            logger.info(f"    {d.strftime('%d/%m/%Y')}  {name}")

    if skipped_future:
        logger.info("-" * 60)
        logger.info(f"  Future days: {len(skipped_future)} (hours won't be filled, but vacation/sick/symbol will be corrected)")

    logger.info("=" * 60)


# ==============================================================================
# Browser Automation
# ==============================================================================

def fix_rtl(text: str) -> str:
    """Reverse Hebrew text so it displays correctly in LTR terminals."""
    if any('\u0590' <= c <= '\u05FF' for c in text):
        return text[::-1]
    return text


def take_screenshot(page: Page, name: str, logger: logging.Logger):
    """Take a screenshot and save it to the screenshots directory."""
    SCREENSHOT_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filepath = SCREENSHOT_DIR / f"{timestamp}_{name}.png"
    page.screenshot(path=str(filepath), full_page=True)
    logger.debug(f"Screenshot saved: {filepath}")


def wait_and_retry(page: Page, selector: str, timeout: int = 10000,
                   retries: int = 3, logger: logging.Logger = None) -> bool:
    """Wait for a selector with retry logic. Returns True if found."""
    for attempt in range(1, retries + 1):
        try:
            page.wait_for_selector(selector, timeout=timeout, state="visible")
            return True
        except PlaywrightTimeoutError:
            if logger:
                logger.warning(f"Selector '{selector}' not found (attempt {attempt}/{retries})")
            if attempt < retries:
                time.sleep(2)
    return False


def dismiss_modal(page: Page, logger: logging.Logger):
    """
    Dismiss any visible modal dialog on the page.
    Hilan uses ASP.NET HModalPopupExtender (id="MPEBehavior") which shows
    a background overlay and a dialog inside an iframe.
    We dismiss it by calling the JS API: $find('MPEBehavior').hide()
    """
    try:
        # Check if any modal background is visible and dismiss it via JS API
        dismissed = page.evaluate("""
            () => {
                const results = [];

                // Check MPEBehavior modal (the main validation dialog)
                const mpeBg = document.getElementById('MPEBehavior_backgroundElement');
                if (mpeBg && mpeBg.style.display !== 'none') {
                    try {
                        const mpe = $find('MPEBehavior');
                        if (mpe && mpe.hide) { mpe.hide(); results.push('MPEBehavior'); }
                    } catch(e) { results.push('MPEBehavior_error:' + e.message); }
                }

                // Check mDialogEx modal (another dialog type on the page)
                const mdBg = document.getElementById('ctl00_mDialogEx_backgroundElement');
                if (mdBg && mdBg.style.display !== 'none') {
                    try {
                        const md = $find('ctl00_mDialogEx');
                        if (md && md.hide) { md.hide(); results.push('mDialogEx'); }
                    } catch(e) { results.push('mDialogEx_error:' + e.message); }
                }

                // Check grid-level dialog
                const gridBg = document.querySelector('[id*="chooseChild_backgroundElement"]');
                if (gridBg && gridBg.style.display !== 'none') {
                    try { gridBg.style.display = 'none'; results.push('gridDialog'); }
                    catch(e) {}
                }

                return results;
            }
        """)

        if dismissed:
            logger.info(f"  [Modal dismissed: {dismissed}]")
            time.sleep(1)

    except Exception as e:
        logger.debug(f"  dismiss_modal error: {e}")


def login(page: Page, username: str, password: str, logger: logging.Logger):
    """
    Log in to the Hilan system.
    Selectors verified from the saved login page HTML.
    """
    logger.info("Navigating to login page...")
    page.goto(HILAN_LOGIN_URL, wait_until="networkidle", timeout=30000)
    take_screenshot(page, "01_login_page", logger)

    # Wait for login form to be ready
    logger.info("Waiting for login form...")
    if not wait_and_retry(page, SEL_USERNAME_INPUT, timeout=15000, logger=logger):
        take_screenshot(page, "01_login_form_not_found", logger)
        raise RuntimeError("Login form not found - could not find username input")

    # Fill in credentials
    logger.info(f"Filling in username: {username}")
    page.fill(SEL_USERNAME_INPUT, username)

    logger.info("Filling in password: ****")
    page.fill(SEL_PASSWORD_INPUT, password)

    take_screenshot(page, "02_credentials_filled", logger)

    # Click login button
    logger.info("Clicking login button...")
    page.click(SEL_LOGIN_BUTTON)

    # Wait for navigation after login
    logger.info("Waiting for login to complete...")
    try:
        # Wait for either the home page or an error message
        page.wait_for_load_state("networkidle", timeout=30000)
        time.sleep(3)  # Extra wait for Angular app to initialize
    except PlaywrightTimeoutError:
        logger.warning("Timeout waiting for post-login navigation")

    take_screenshot(page, "03_after_login", logger)

    # Check if login was successful by looking at the URL
    current_url = page.url
    logger.info(f"Current URL after login: {current_url}")

    if "login" in current_url.lower():
        # Check for error message
        error_el = page.query_selector(".error, .h-centered-field.error, [role='alert']")
        error_text = error_el.inner_text() if error_el else "Unknown error"
        take_screenshot(page, "03_login_failed", logger)
        raise RuntimeError(f"Login failed! Error: {fix_rtl(error_text)}")

    logger.info("Login successful!")


def navigate_to_attendance(page: Page, logger: logging.Logger):
    """Navigate to the attendance/calendar page."""
    logger.info("Navigating to attendance page...")
    page.goto(HILAN_ATTENDANCE_URL, wait_until="networkidle", timeout=30000)
    time.sleep(3)  # Wait for ASP.NET page to fully initialize

    take_screenshot(page, "04_attendance_page", logger)
    logger.info(f"Attendance page URL: {page.url}")


def select_month_in_calendar(page: Page, year: int, month: int, logger: logging.Logger):
    """
    Navigate to the correct month in the Hilan calendar.
    The calendar has a month dropdown (BulletedList) with items like "February 2026".
    Clicking an item triggers ChangeMonthDdl() which posts back.
    """
    logger.info(f"Checking if calendar shows {month}/{year}...")

    # Check current month shown in the calendar header
    month_label = page.query_selector("#ctl00_mp_calendar_monthChanged")
    if month_label:
        current_text = month_label.inner_text().strip()
        target_text = f"{calendar.month_name[month]} {year}"
        logger.info(f"Calendar currently shows: '{current_text}', target: '{target_text}'")

        if target_text.lower() == current_text.lower():
            logger.info("Already on the correct month!")
            return

    # Need to switch month - click the month label to open the dropdown
    logger.info("Opening month selector dropdown...")
    dropdown_container = page.query_selector(".SelectedBulletedListItem")
    if dropdown_container:
        dropdown_container.click()
        time.sleep(0.5)

        # Find the target month item in the dropdown list
        target_date_str = f"01/{str(month).zfill(2)}/{year}"
        month_item = page.query_selector(f"li[itemvalue='{target_date_str}']")
        if month_item:
            logger.info(f"Clicking month: {month_item.inner_text().strip()}")
            month_item.click()
            # Wait for postback to complete
            page.wait_for_load_state("networkidle", timeout=30000)
            time.sleep(3)
            logger.info("Month changed successfully!")
        else:
            logger.warning(f"Could not find month item for {target_date_str}")
    else:
        logger.warning("Month selector dropdown not found")

    take_screenshot(page, "05_month_selected", logger)


def select_workdays_in_calendar(page: Page, workdays: list[date], logger: logging.Logger):
    """
    Select all days in the month by clicking the ">>" (select all) button.
    This is more reliable than clicking individual days (which can miss days
    due to ASP.NET postback timing issues with many clicks).
    The fill logic already filters to only fill target workdays.
    """
    logger.info(f"Selecting all days in month (>> select all)...")

    # Click ">>" to select all days of the month
    select_all_btn = page.query_selector("th.dayFirstHeaderStyle")
    if select_all_btn:
        select_all_btn.click()
        try:
            page.wait_for_load_state("networkidle", timeout=30000)
        except PlaywrightTimeoutError:
            pass
        time.sleep(2)

        # >> toggles selection - if today is in the viewed month, it was
        # already selected, so >> deselected it. Re-click to re-select.
        # IMPORTANT: only do this when viewing the CURRENT month, otherwise
        # we'd accidentally deselect a random day in a different month.
        today = date.today()
        viewed_month = workdays[0].month if workdays else None
        viewed_year = workdays[0].year if workdays else None
        is_current_month = (viewed_month == today.month and viewed_year == today.year)

        if is_current_month:
            today_cell = page.query_selector(f"td.currentDay[aria-label='{today.day}'], td.CSD[aria-label='{today.day}']")
            if today_cell:
                today_cell.click()
                try:
                    page.wait_for_load_state("networkidle", timeout=10000)
                except PlaywrightTimeoutError:
                    pass
                time.sleep(1)
                logger.info(f"All days selected via >> button (re-selected today: {today.day}).")
            else:
                logger.info("All days selected via >> button.")
        else:
            logger.info(f"All days selected via >> button (not current month, no re-select needed).")
    else:
        logger.warning(">> button not found, falling back to individual day clicks...")
        # Fallback: click individual days
        for workday in workdays:
            day_num = workday.day
            day_cell = page.query_selector(
                f"td[aria-label='{day_num}'][onclick*='_dSD']"
            )
            if day_cell:
                day_cell.click()
                try:
                    page.wait_for_load_state("networkidle", timeout=15000)
                except PlaywrightTimeoutError:
                    pass
                time.sleep(1)
            else:
                logger.warning(f"  Calendar cell for day {day_num} not found!")

    # Click "Days selected" button to show the grid with selected days
    logger.info("Clicking 'Days selected' button...")
    days_selected_btn = page.query_selector("#ctl00_mp_RefreshSelectedDays")
    if not days_selected_btn:
        days_selected_btn = page.query_selector("input[value='Days selected']")
    if days_selected_btn:
        days_selected_btn.click()
        try:
            page.wait_for_load_state("networkidle", timeout=30000)
        except PlaywrightTimeoutError:
            pass
        time.sleep(3)
        logger.info("'Days selected' clicked - grid should now show all selected workdays.")
    else:
        logger.warning("'Days selected' button not found!")

    take_screenshot(page, "06_workdays_selected", logger)
    logger.info(f"Selected {len(workdays)} days. Grid should now show their rows.")


def get_grid_rows_info(page: Page, logger: logging.Logger) -> list[dict]:
    """
    Read all rows from the reports grid and return their info.
    Each row has: row_index, date_text, input IDs for entry/exit/project, current values.
    """
    rows_info = page.evaluate("""
        () => {
            const rows = [];

            // Strategy: find all date cells (ReportDate) and work outwards to find
            // the corresponding entry/exit/project inputs using the row index.
            // Date spans have IDs like: ..._cellOf_ReportDate_row_N_ctl00
            const dateSpans = document.querySelectorAll('span[id*="ReportDate_row_"]');

            for (const dateSpan of dateSpans) {
                const spanId = dateSpan.id;
                // Extract row index: ..._ReportDate_row_N_ctl00  or  ..._ReportDate_row_N_...
                const match = spanId.match(/_row_(\\d+)/);
                if (!match) continue;
                const rowIndex = parseInt(match[1]);

                const dateText = (dateSpan.getAttribute('title') || dateSpan.innerText || '').trim();

                // Find ManualEntry input (search in whole document by row index)
                const entryInput = document.querySelector(
                    `input[id*="ManualEntry_EmployeeReports_row_${rowIndex}_"]`
                );
                // Find ManualExit input
                const exitInput = document.querySelector(
                    `input[id*="ManualExit_EmployeeReports_row_${rowIndex}_"]`
                );

                // Find Project autocomplete text input
                const projectTd = document.querySelector(
                    `td[id*="Project"][id*="_EmployeeReports_row_${rowIndex}_"]`
                );
                const projectInput = projectTd
                    ? projectTd.querySelector('input[type="text"]')
                    : null;
                // The hidden input stores the actual project value
                const projectHidden = document.querySelector(
                    `input[id*="ProjectForView_EmployeeReports_row_${rowIndex}_"][id*="AutoCompleteExtender_value"]`
                );

                // Find Reporting Type (Symbol) dropdown
                const symbolSelect = document.querySelector(
                    `select[id*="Symbol"][id*="_EmployeeReports_row_${rowIndex}_"]`
                );

                rows.push({
                    rowIndex: rowIndex,
                    dateText: dateText,
                    hasEntryInput: !!entryInput,
                    hasExitInput: !!exitInput,
                    entryInputId: entryInput ? entryInput.id : null,
                    exitInputId: exitInput ? exitInput.id : null,
                    currentEntry: entryInput ? entryInput.value : '',
                    currentExit: exitInput ? exitInput.value : '',
                    hasProjectInput: !!projectInput,
                    projectInputName: projectInput ? projectInput.name : null,
                    projectHiddenId: projectHidden ? projectHidden.id : null,
                    currentProject: projectHidden ? projectHidden.value : '',
                    currentProjectText: projectInput ? projectInput.value : '',
                    hasSymbolSelect: !!symbolSelect,
                    symbolSelectId: symbolSelect ? symbolSelect.id : null,
                    currentSymbol: symbolSelect ? symbolSelect.value : '',
                });
            }
            return rows;
        }
    """)

    logger.info(f"Found {len(rows_info)} day rows in the reports grid:")
    for row in rows_info:
        logger.info(f"  Row {row['rowIndex']}: date='{row['dateText']}' "
                     f"entry={row['currentEntry'] or '(empty)'} "
                     f"exit={row['currentExit'] or '(empty)'} "
                     f"inputs={'YES' if row['hasEntryInput'] else 'NO'} "
                     f"project={'YES' if row['hasProjectInput'] else 'NO'} "
                     f"symbol={row.get('currentSymbol', '?')}|{'YES' if row.get('hasSymbolSelect') else 'NO'}")

    return rows_info


def fill_project_field(page: Page, row: dict, project: str, logger: logging.Logger) -> bool:
    """
    Fill the project autocomplete field for a given row.
    Uses Playwright to type in the field and select from the autocomplete dropdown,
    because ASP.NET AutoCompleteExtender needs real keyboard input to trigger its
    AJAX search and properly set the hidden value.
    Returns True on success.
    """
    if not row.get("hasProjectInput") or not row.get("projectInputName"):
        logger.warning(f"    Project input not found for row {row['rowIndex']}")
        return False

    row_index = row['rowIndex']
    project_name = row['projectInputName']

    try:
        project_selector = f"input[name='{project_name}']"

        # Step 0: Close any open autocomplete dropdowns AND iframes from previous rows
        # These can block clicks on the current row's input
        page.evaluate("""
            () => {
                // Hide autocomplete dropdowns
                document.querySelectorAll('ul[id*="AutoCompleteExtender_completionListElem"]').forEach(ul => {
                    ul.style.display = 'none';
                    ul.style.visibility = 'hidden';
                });
                // Hide autocomplete iframes that intercept pointer events
                document.querySelectorAll('iframe[src*="javascript"]').forEach(f => {
                    f.style.display = 'none';
                });
            }
        """)
        time.sleep(0.2)

        # Step 1: Clear watermark via JS, then click with Playwright
        page.evaluate("""
            (params) => {
                const projectTd = document.querySelector(
                    `td[id*="Project"][id*="_EmployeeReports_row_${params.rowIndex}_"]`
                );
                if (!projectTd) return;
                const textInput = projectTd.querySelector('input[type="text"]');
                if (!textInput) return;
                textInput.classList.remove('Watermarked');
                textInput.value = '';
            }
        """, {"rowIndex": row_index})
        time.sleep(0.2)

        # Double-click sequence: first click activates the ASP.NET control (doControlClick),
        # second click ensures the autocomplete is ready. The first row often needs this
        # "warm-up" because the autocomplete AJAX handler isn't initialized until first focus.
        page.click(project_selector, force=True)
        time.sleep(0.5)
        # Click again to ensure autocomplete is fully activated
        page.click(project_selector, force=True)
        time.sleep(0.3)

        # Step 2: Type just the project code (number before " - ") for precise search
        project_code = project.split(" - ")[0].strip() if " - " in project else project
        project_selector = f"input[name='{project_name}']"

        # Try up to 3 attempts (autocomplete may need more time, especially first row)
        for attempt in range(3):
            if attempt > 0:
                # Retry: clear, re-click, and retype
                logger.info(f"    Project: retrying autocomplete (attempt {attempt + 1})...")
                page.evaluate("""
                    (params) => {
                        const projectTd = document.querySelector(
                            `td[id*="Project"][id*="_EmployeeReports_row_${params.rowIndex}_"]`
                        );
                        if (!projectTd) return;
                        const textInput = projectTd.querySelector('input[type="text"]');
                        if (!textInput) return;
                        textInput.classList.remove('Watermarked');
                        textInput.value = '';
                    }
                """, {"rowIndex": row_index})
                time.sleep(0.3)
                page.click(project_selector, force=True)
                time.sleep(0.5)
                page.type(project_selector, project_code, delay=80)
            else:
                page.type(project_selector, project_code, delay=80)

            time.sleep(3)  # Wait for autocomplete AJAX

            # Find the autocomplete dropdown - first try row-specific, then ANY visible one
            autocomplete_list = page.query_selector(
                f"ul[id*='ProjectForView_EmployeeReports_row_{row_index}_'][id*='completionListElem']"
            )
            if not (autocomplete_list and autocomplete_list.is_visible()):
                # Fallback: find ANY visible autocomplete list (ASP.NET may reuse them)
                all_lists = page.query_selector_all("ul[id*='AutoCompleteExtender_completionListElem']")
                for lst in all_lists:
                    if lst.is_visible():
                        autocomplete_list = lst
                        break

            if autocomplete_list and autocomplete_list.is_visible():
                items = autocomplete_list.query_selector_all("li")
                matched_item = None
                for item in items:
                    item_text = item.inner_text().strip()
                    if item_text.startswith(project_code):
                        matched_item = item
                        break
                if not matched_item:
                    for item in items:
                        item_text = item.inner_text().strip()
                        if project_code in item_text:
                            matched_item = item
                            break
                if not matched_item and items:
                    matched_item = items[0]

                if matched_item:
                    item_text = matched_item.inner_text().strip()
                    matched_item.click()
                    time.sleep(0.5)
                    logger.info(f"    Project selected: {item_text}")
                    return True

        # Final fallback: set value via JS (may not be recognized by ASP.NET but better than nothing)
        logger.warning(f"    Project: autocomplete failed, setting via JS fallback")
        page.evaluate("""
            (params) => {
                const projectTd = document.querySelector(
                    `td[id*="Project"][id*="_EmployeeReports_row_${params.rowIndex}_"]`
                );
                if (!projectTd) return;
                const textInput = projectTd.querySelector('input[type="text"]');
                const hiddenInput = document.querySelector(
                    `input[id*="ProjectForView_EmployeeReports_row_${params.rowIndex}_"][id*="AutoCompleteExtender_value"]`
                );
                if (textInput) {
                    textInput.classList.remove('Watermarked');
                    textInput.value = params.project;
                    textInput.setAttribute('sv', params.project);
                    textInput.dispatchEvent(new Event('change', {bubbles: true}));
                }
                if (hiddenInput) { hiddenInput.value = params.project; }
            }
        """, {"rowIndex": row_index, "project": project})
        time.sleep(0.5)
        # Close any open autocomplete dropdowns
        page.evaluate('document.querySelectorAll("ul[id*=AutoCompleteExtender_completionListElem]").forEach(ul => { ul.style.display = "none"; })')
        logger.info(f"    Project set via JS: {project}")
        return False  # Return False so this row gets retried after autocomplete is warmed up

    except Exception as e:
        logger.warning(f"    Project: Error filling project: {e}")
        # Always close any stale autocomplete dropdowns on error
        try:
            page.evaluate('document.querySelectorAll("ul[id*=AutoCompleteExtender_completionListElem]").forEach(ul => { ul.style.display = "none"; })')
        except Exception:
            pass
        return False


def fill_all_hours(page: Page, workdays: list[date], entry_time: str,
                   exit_time: str, project: str | None,
                   present_weekdays: set, present_dates: set,
                   vacation_days: set,
                   sick_days: set, sick_file: str | None,
                   logger: logging.Logger) -> tuple[int, int]:
    """
    Fill hours, project, reporting type, vacation, and sick days for all workdays.
    The grid must already show all days (call select_all_workdays_in_calendar first).
    Returns (success_count, failure_count).
    """
    success_count = 0
    failure_count = 0
    skipped_count = 0

    # Build a set of workday date strings for quick lookup (format: "DD/MM")
    workday_dates = set()
    for wd in workdays:
        workday_dates.add(f"{wd.day:02d}/{wd.month:02d}")

    # All workdays (including future) are already in the list from get_workdays()

    # Get all rows from the grid
    rows_info = get_grid_rows_info(page, logger)

    if not rows_info:
        logger.error("No rows found in the reports grid!")
        logger.error("The page might not have loaded correctly, or selectors need updating.")
        take_screenshot(page, "error_no_rows", logger)

        # Dump page HTML for debugging
        html_path = SCREENSHOT_DIR / "attendance_page_debug.html"
        SCREENSHOT_DIR.mkdir(exist_ok=True)
        html_content = page.content()
        html_path.write_text(html_content, encoding="utf-8")
        logger.info(f"Page HTML saved to: {html_path}")
        return 0, len(workdays)

    # Save page HTML for debugging if no rows matched
    take_screenshot(page, "07_grid_before_fill", logger)

    logger.info(f"\nFilling hours for workdays...")
    if project:
        logger.info(f"Project: {project}")
    if present_dates:
        logger.info(f"Present dates (specific): {sorted(present_dates)}")
    if vacation_days:
        logger.info(f"Vacation days: {sorted(vacation_days)}")
    if sick_days:
        sick_type = "sick day declaration" if len(sick_days) <= 2 else f"sick (file: {sick_file})"
        logger.info(f"Sick days: {sorted(sick_days)} ({sick_type})")
    logger.info(f"Looking for dates: {workday_dates}")
    logger.info("-" * 60)

    rows_needing_project_retry = []

    # ==========================================
    # PASS 1: Postback operations (vacation, sick, delete, symbol changes)
    # These trigger ASP.NET postbacks that reload the grid.
    # Must be done FIRST so they don't wipe project values set later.
    #
    # IMPORTANT: Each postback reloads the grid with new row indices.
    # We must re-read the grid after EACH postback and restart the scan,
    # otherwise subsequent operations use stale row indices and fail silently.
    # We use a while loop that keeps scanning until no more changes are needed.
    # ==========================================
    logger.info("--- Pass 1: Postback operations (vacation, sick, delete, symbol changes) ---")
    postback_occurred = False

    for row in rows_info:
        date_text = row["dateText"]  # e.g., "01/02 Sun"

        # Extract DD/MM from the date text
        date_parts = date_text.split(" ")
        if not date_parts:
            continue
        date_ddmm = date_parts[0]  # "01/02"

        # Check if this is a Fri/Sat row
        day_name = date_parts[1] if len(date_parts) > 1 else ""
        if day_name in ("Fri", "Sat"):
            current_sym = row.get("currentSymbol", "")
            has_hours = bool(row.get("currentEntry", "").strip() or row.get("currentExit", "").strip())
            harmless_symbols = {"", "15"}  # empty (Select) or w.home
            # Only delete Fri/Sat rows that have actual data (hours) or non-trivial
            # symbols (vacation/sick/presence). Empty rows or rows with just
            # w.home/Select and no hours are harmless - skip them to save time.
            if current_sym and (has_hours or current_sym not in harmless_symbols):
                # Delete this incorrect Fri/Sat row via doControlClick (ASP.NET handler)
                try:
                    page.evaluate("""
                        (params) => {
                            const deleteSpan = document.querySelector(
                                `span[id*='SysColumn_Delete_EmployeeReports_row_${params.rowIndex}_'] span.fh-garbage`
                            );
                            if (deleteSpan && typeof doControlClick === 'function') {
                                doControlClick({type:'click'}, deleteSpan);
                            }
                        }
                    """, {"rowIndex": row['rowIndex']})
                    time.sleep(2)

                    # Click "OK"/"Yes" inside the confirmation dialog iframe
                    popup = page.query_selector("#ctl00_theBasePanel")
                    if popup:
                        iframe = popup.query_selector("iframe")
                        if iframe:
                            frame = iframe.content_frame()
                            if frame:
                                # Try clicking OK/Yes button inside the iframe
                                ok_btn = frame.query_selector("input[value='OK'], input[value='Yes'], button:has-text('OK'), button:has-text('Yes'), a:has-text('OK')")
                                if ok_btn:
                                    ok_btn.click()
                                    logger.info(f"  {date_text}: Confirmed delete in dialog")
                                    time.sleep(2)

                    try:
                        page.wait_for_load_state("networkidle", timeout=15000)
                    except PlaywrightTimeoutError:
                        pass
                    time.sleep(1)
                    dismiss_modal(page, logger)
                    logger.info(f"  {date_text}: Deleted incorrect Fri/Sat row (was {current_sym})")
                    success_count += 1
                    postback_occurred = True
                    break  # Re-read grid after postback (row indices changed)
                except Exception as e:
                    logger.warning(f"  {date_text}: Error deleting Fri/Sat row: {e}")
                    dismiss_modal(page, logger)
            else:
                logger.debug(f"  {date_text}: Fri/Sat with no data, skipping (symbol={current_sym})")
            continue

        # Check if this row is a workday we want to fill
        if date_ddmm not in workday_dates:
            continue

        day_num = int(date_ddmm.split("/")[0])

        # Handle vacation days: set symbol to "vacation" (2), clear hours/project
        is_vacation_day = day_num in vacation_days
        if is_vacation_day:
            current_symbol = row.get("currentSymbol", "")
            has_hours = bool(row.get("currentEntry", "").strip() or row.get("currentExit", "").strip())
            if current_symbol == "2" and not has_hours:
                logger.info(f"  {date_text}: Already set to vacation, skipping")
                skipped_count += 1
                continue
            if row.get("hasSymbolSelect") and row.get("symbolSelectId"):
                try:
                    dismiss_modal(page, logger)

                    # Clear hours AND project (vacation = absence, no hours/project needed)
                    entry_id = row.get('entryInputId')
                    exit_id = row.get('exitInputId')
                    row_index = row['rowIndex']
                    page.evaluate("""
                        (params) => {
                            // Clear hours
                            const entryEl = document.getElementById(params.entryId);
                            const exitEl = document.getElementById(params.exitId);
                            if (entryEl) { entryEl.value = ''; }
                            if (exitEl) { exitEl.value = ''; }

                            // Clear project (both visible text and hidden value)
                            const projectTd = document.querySelector(
                                `td[id*="Project"][id*="_EmployeeReports_row_${params.rowIndex}_"]`
                            );
                            if (projectTd) {
                                const textInput = projectTd.querySelector('input[type="text"]');
                                if (textInput) {
                                    textInput.value = '';
                                    textInput.setAttribute('sv', '');
                                }
                            }
                            const hiddenInput = document.querySelector(
                                `input[id*="ProjectForView_EmployeeReports_row_${params.rowIndex}_"][id*="AutoCompleteExtender_value"]`
                            );
                            if (hiddenInput) { hiddenInput.value = ''; }
                        }
                    """, {"entryId": entry_id or "", "exitId": exit_id or "", "rowIndex": row_index})

                    # Set symbol to vacation using Playwright (not JS) to trigger
                    # the real ASP.NET onSelectionChanged handler
                    select_selector = f"select[id='{row['symbolSelectId']}']"
                    page.select_option(select_selector, "2")
                    logger.info(f"  {date_text}: Set to VACATION (hours cleared)")
                    try:
                        page.wait_for_load_state("networkidle", timeout=10000)
                    except PlaywrightTimeoutError:
                        pass
                    time.sleep(1)
                    dismiss_modal(page, logger)
                    success_count += 1
                    postback_occurred = True
                    break  # Re-read grid after postback (row indices changed)
                except Exception as e:
                    logger.error(f"  {date_text}: Error setting vacation: {e}")
                    failure_count += 1
            else:
                logger.warning(f"  {date_text}: No symbol dropdown for vacation")
                failure_count += 1
            continue

        # Handle sick days
        is_sick_day = day_num in sick_days
        if is_sick_day:
            # 1-2 days: "sick day declaration" (value=5), 3+: "sick" (value=6)
            sick_symbol = "5" if len(sick_days) <= 2 else "6"
            sick_label = "sick day declaration" if sick_symbol == "5" else "sick"
            current_symbol = row.get("currentSymbol", "")
            has_hours = bool(row.get("currentEntry", "").strip() or row.get("currentExit", "").strip())
            if current_symbol == sick_symbol and not has_hours:
                logger.info(f"  {date_text}: Already set to {sick_label}, skipping")
                skipped_count += 1
                continue
            if row.get("hasSymbolSelect") and row.get("symbolSelectId"):
                try:
                    dismiss_modal(page, logger)

                    # Clear hours and project
                    entry_id = row.get('entryInputId')
                    exit_id = row.get('exitInputId')
                    row_index = row['rowIndex']
                    page.evaluate("""
                        (params) => {
                            const entryEl = document.getElementById(params.entryId);
                            const exitEl = document.getElementById(params.exitId);
                            if (entryEl) { entryEl.value = ''; }
                            if (exitEl) { exitEl.value = ''; }
                            const projectTd = document.querySelector(
                                `td[id*="Project"][id*="_EmployeeReports_row_${params.rowIndex}_"]`
                            );
                            if (projectTd) {
                                const textInput = projectTd.querySelector('input[type="text"]');
                                if (textInput) { textInput.value = ''; textInput.setAttribute('sv', ''); }
                            }
                            const hiddenInput = document.querySelector(
                                `input[id*="ProjectForView_EmployeeReports_row_${params.rowIndex}_"][id*="AutoCompleteExtender_value"]`
                            );
                            if (hiddenInput) { hiddenInput.value = ''; }
                        }
                    """, {"entryId": entry_id or "", "exitId": exit_id or "", "rowIndex": row_index})

                    # Set symbol to sick/sick day declaration
                    select_selector = f"select[id='{row['symbolSelectId']}']"
                    page.select_option(select_selector, sick_symbol)
                    logger.info(f"  {date_text}: Set to {sick_label.upper()}")
                    try:
                        page.wait_for_load_state("networkidle", timeout=10000)
                    except PlaywrightTimeoutError:
                        pass
                    time.sleep(1)
                    dismiss_modal(page, logger)

                    # Upload sick certificate file if needed (3+ days, symbol=6)
                    if sick_symbol == "6" and sick_file:
                        try:
                            # Find the Upload link for this row
                            upload_span = page.query_selector(
                                f"span[id*='File_EmployeeReports_row_{row_index}_'][id*='Attach']"
                            )
                            if upload_span and upload_span.is_visible():
                                # Use Playwright file chooser to handle the upload dialog
                                with page.expect_file_chooser(timeout=10000) as fc_info:
                                    upload_span.click(force=True)
                                file_chooser = fc_info.value
                                file_chooser.set_files(sick_file)
                                time.sleep(2)
                                try:
                                    page.wait_for_load_state("networkidle", timeout=15000)
                                except PlaywrightTimeoutError:
                                    pass
                                dismiss_modal(page, logger)
                                logger.info(f"    Sick certificate uploaded: {sick_file}")
                            else:
                                logger.warning(f"    Upload link not found for row {row_index}")
                        except Exception as e:
                            logger.warning(f"    Error uploading sick file: {e}")

                    success_count += 1
                    postback_occurred = True
                    break  # Re-read grid after postback (row indices changed)
                except Exception as e:
                    logger.error(f"  {date_text}: Error setting sick: {e}")
                    failure_count += 1
            else:
                logger.warning(f"  {date_text}: No symbol dropdown for sick")
                failure_count += 1
            continue

        # Handle symbol corrections that trigger postbacks
        # This runs for ALL workdays, not just when present-days is set.
        if row.get("hasSymbolSelect"):
            current_symbol = row.get("currentSymbol", "")
            absence_symbols = {"2", "5", "6"}  # vacation, sick day decl, sick

            # Determine expected symbol for this day
            matching_wd = None
            for wd in workdays:
                if f"{wd.day:02d}/{wd.month:02d}" == date_ddmm:
                    matching_wd = wd
                    break

            if matching_wd:
                # A day is "present" if it matches a recurring weekday OR a specific date
                is_present_day = (matching_wd.weekday() in present_weekdays or
                                  matching_wd.day in present_dates)
                expected_symbol = "0" if is_present_day else "15"

                # STEP 1: Delete absence rows that shouldn't be absence anymore
                # If a day currently has vacation/sick but is NOT in the current
                # vacation/sick sets, it must be deleted (ASP.NET can't change
                # directly from absence to work type).
                if current_symbol in absence_symbols:
                    try:
                        page.evaluate("""
                            (params) => {
                                const deleteSpan = document.querySelector(
                                    `span[id*='SysColumn_Delete_EmployeeReports_row_${params.rowIndex}_'] span.fh-garbage`
                                );
                                if (deleteSpan && typeof doControlClick === 'function') {
                                    doControlClick({type:'click'}, deleteSpan);
                                }
                            }
                        """, {"rowIndex": row['rowIndex']})
                        time.sleep(2)
                        popup = page.query_selector("#ctl00_theBasePanel")
                        if popup:
                            iframe = popup.query_selector("iframe")
                            if iframe:
                                frame = iframe.content_frame()
                                if frame:
                                    ok_btn = frame.query_selector("input[value='OK'], input[value='Yes'], button:has-text('OK')")
                                    if ok_btn:
                                        ok_btn.click()
                                        time.sleep(2)
                        try:
                            page.wait_for_load_state("networkidle", timeout=15000)
                        except PlaywrightTimeoutError:
                            pass
                        time.sleep(1)
                        dismiss_modal(page, logger)
                        logger.info(f"  {date_text}: Deleted absence row (was {current_symbol})")
                        success_count += 1
                        postback_occurred = True
                        break  # Re-read grid after postback (row indices changed)
                    except Exception as e:
                        logger.warning(f"  {date_text}: Error deleting absence row: {e}")
                        dismiss_modal(page, logger)

                # STEP 2: Change symbol between presence and w.home if needed
                elif current_symbol and current_symbol != expected_symbol:
                    try:
                        dismiss_modal(page, logger)
                        symbol_label = "presence" if expected_symbol == "0" else "w.home"
                        select_selector = f"select[id='{row['symbolSelectId']}']"
                        page.select_option(select_selector, expected_symbol)
                        logger.info(f"  {date_text}: Reporting Type set to: {symbol_label}")
                        try:
                            page.wait_for_load_state("networkidle", timeout=10000)
                        except PlaywrightTimeoutError:
                            pass
                        time.sleep(1)
                        dismiss_modal(page, logger)
                        success_count += 1
                        postback_occurred = True
                        break  # Re-read grid after postback (row indices changed)
                    except Exception as e:
                        logger.warning(f"  {date_text}: Error changing symbol: {e}")
                        dismiss_modal(page, logger)

    # ==========================================
    # RE-READ & RESCAN: After postbacks, row indices changed.
    # Keep re-reading and re-scanning until no more postbacks are needed.
    # ==========================================
    _rescan_count = 0
    while postback_occurred and _rescan_count < 50:
        _rescan_count += 1
        logger.info("")
        logger.info(f"--- Re-reading grid after postback operations (rescan {_rescan_count}) ---")
        rows_info = get_grid_rows_info(page, logger)
        if not rows_info:
            logger.error("No rows found after re-read!")
            take_screenshot(page, "error_no_rows_after_reread", logger)
            return success_count, failure_count

        _did_postback = False
        for row in rows_info:
            date_text = row["dateText"]
            date_parts = date_text.split(" ")
            if not date_parts:
                continue
            date_ddmm = date_parts[0]
            day_name = date_parts[1] if len(date_parts) > 1 else ""

            # Skip Fri/Sat (already handled in first scan)
            if day_name in ("Fri", "Sat"):
                continue

            if date_ddmm not in workday_dates:
                continue

            day_num = int(date_ddmm.split("/")[0])

            # Check vacation days that still need setting
            if day_num in vacation_days:
                current_symbol = row.get("currentSymbol", "")
                has_hours = bool(row.get("currentEntry", "").strip() or row.get("currentExit", "").strip())
                if current_symbol == "2" and not has_hours:
                    continue  # Already correct
                if row.get("hasSymbolSelect") and row.get("symbolSelectId"):
                    try:
                        dismiss_modal(page, logger)
                        entry_id = row.get('entryInputId')
                        exit_id = row.get('exitInputId')
                        row_index = row['rowIndex']
                        page.evaluate("""
                            (params) => {
                                const e = document.getElementById(params.entryId);
                                const x = document.getElementById(params.exitId);
                                if (e) { e.value = ''; } if (x) { x.value = ''; }
                            }
                        """, {"entryId": entry_id or "", "exitId": exit_id or ""})
                        select_selector = f"select[id='{row['symbolSelectId']}']"
                        page.select_option(select_selector, "2")
                        logger.info(f"  {date_text}: Set to VACATION (rescan)")
                        try:
                            page.wait_for_load_state("networkidle", timeout=10000)
                        except PlaywrightTimeoutError:
                            pass
                        time.sleep(1)
                        dismiss_modal(page, logger)
                        success_count += 1
                        _did_postback = True
                        break  # Re-read after postback
                    except Exception as e:
                        logger.error(f"  {date_text}: Error setting vacation (rescan): {e}")
                continue

            # Check sick days that still need setting
            if day_num in sick_days:
                sick_symbol = "5" if len(sick_days) <= 2 else "6"
                current_symbol = row.get("currentSymbol", "")
                has_hours = bool(row.get("currentEntry", "").strip() or row.get("currentExit", "").strip())
                if current_symbol == sick_symbol and not has_hours:
                    continue  # Already correct
                if row.get("hasSymbolSelect") and row.get("symbolSelectId"):
                    try:
                        dismiss_modal(page, logger)
                        entry_id = row.get('entryInputId')
                        exit_id = row.get('exitInputId')
                        page.evaluate("""
                            (params) => {
                                const e = document.getElementById(params.entryId);
                                const x = document.getElementById(params.exitId);
                                if (e) { e.value = ''; } if (x) { x.value = ''; }
                            }
                        """, {"entryId": entry_id or "", "exitId": exit_id or ""})
                        select_selector = f"select[id='{row['symbolSelectId']}']"
                        page.select_option(select_selector, sick_symbol)
                        logger.info(f"  {date_text}: Set to SICK (rescan)")
                        try:
                            page.wait_for_load_state("networkidle", timeout=10000)
                        except PlaywrightTimeoutError:
                            pass
                        time.sleep(1)
                        dismiss_modal(page, logger)
                        success_count += 1
                        _did_postback = True
                        break
                    except Exception as e:
                        logger.error(f"  {date_text}: Error setting sick (rescan): {e}")
                continue

            # Check absence symbols that need deletion
            if row.get("hasSymbolSelect"):
                current_symbol = row.get("currentSymbol", "")
                absence_symbols = {"2", "5", "6"}
                if current_symbol in absence_symbols:
                    try:
                        page.evaluate("""
                            (params) => {
                                const d = document.querySelector(
                                    `span[id*='SysColumn_Delete_EmployeeReports_row_${params.rowIndex}_'] span.fh-garbage`
                                );
                                if (d && typeof doControlClick === 'function') { doControlClick({type:'click'}, d); }
                            }
                        """, {"rowIndex": row['rowIndex']})
                        time.sleep(2)
                        popup = page.query_selector("#ctl00_theBasePanel")
                        if popup:
                            iframe = popup.query_selector("iframe")
                            if iframe:
                                frame = iframe.content_frame()
                                if frame:
                                    ok_btn = frame.query_selector("input[value='OK'], input[value='Yes'], button:has-text('OK')")
                                    if ok_btn:
                                        ok_btn.click()
                                        time.sleep(2)
                        try:
                            page.wait_for_load_state("networkidle", timeout=15000)
                        except PlaywrightTimeoutError:
                            pass
                        time.sleep(1)
                        dismiss_modal(page, logger)
                        logger.info(f"  {date_text}: Deleted absence row (rescan, was {current_symbol})")
                        success_count += 1
                        _did_postback = True
                        break
                    except Exception as e:
                        logger.warning(f"  {date_text}: Error deleting absence (rescan): {e}")
                        dismiss_modal(page, logger)

                # Check symbol changes (presence <-> w.home)
                matching_wd = None
                for wd in workdays:
                    if f"{wd.day:02d}/{wd.month:02d}" == date_ddmm:
                        matching_wd = wd
                        break
                if matching_wd:
                    is_present_day = (matching_wd.weekday() in present_weekdays or
                                      matching_wd.day in present_dates)
                    expected_symbol = "0" if is_present_day else "15"
                    if current_symbol and current_symbol != expected_symbol:
                        try:
                            dismiss_modal(page, logger)
                            symbol_label = "presence" if expected_symbol == "0" else "w.home"
                            select_selector = f"select[id='{row['symbolSelectId']}']"
                            page.select_option(select_selector, expected_symbol)
                            logger.info(f"  {date_text}: Reporting Type set to: {symbol_label} (rescan)")
                            try:
                                page.wait_for_load_state("networkidle", timeout=10000)
                            except PlaywrightTimeoutError:
                                pass
                            time.sleep(1)
                            dismiss_modal(page, logger)
                            success_count += 1
                            _did_postback = True
                            break
                        except Exception as e:
                            logger.warning(f"  {date_text}: Error changing symbol (rescan): {e}")
                            dismiss_modal(page, logger)

        if not _did_postback:
            break  # No more changes needed

    # Final re-read for Pass 2
    if postback_occurred:
        logger.info("")
        logger.info("--- Final grid re-read for Pass 2 ---")
        rows_info = get_grid_rows_info(page, logger)
        if not rows_info:
            logger.error("No rows found after final re-read!")
            take_screenshot(page, "error_no_rows_final", logger)
            return success_count, failure_count

    # ==========================================
    # PASS 2: Fill hours and projects (JS-only, no postbacks)
    # Now that all postback operations are done, the grid is stable.
    # ==========================================
    logger.info("")
    logger.info("--- Pass 2: Fill hours and projects ---")

    for row in rows_info:
        date_text = row["dateText"]
        date_parts = date_text.split(" ")
        if not date_parts:
            continue
        date_ddmm = date_parts[0]

        # Skip Fri/Sat
        day_name = date_parts[1] if len(date_parts) > 1 else ""
        if day_name in ("Fri", "Sat"):
            continue

        # Skip non-workdays
        if date_ddmm not in workday_dates:
            continue

        day_num = int(date_ddmm.split("/")[0])

        # Skip vacation and sick days (already handled in pass 1)
        if day_num in vacation_days or day_num in sick_days:
            continue

        # CRITICAL: If the page/browser has closed, stop immediately
        if page.is_closed():
            logger.error(f"  Page is closed! Stopping fill loop at row {date_text}.")
            break

        # Check if inputs exist for this row
        if not row["hasEntryInput"] or not row["hasExitInput"]:
            logger.warning(f"  {date_text}: No entry/exit inputs (may be holiday/weekend row)")
            skipped_count += 1
            continue

        # Check if this is a future day
        is_future_day = False
        try:
            dd = int(date_ddmm.split("/")[0])
            mm = int(date_ddmm.split("/")[1])
            yr = workdays[0].year if workdays else date.today().year
            is_future_day = date(yr, mm, dd) > date.today()
        except (ValueError, IndexError):
            pass

        hours_correct = (row["currentEntry"].strip() == entry_time and
                         row["currentExit"].strip() == exit_time)
        hours_need_fill = not hours_correct and not is_future_day

        # Check if project needs filling
        # Compare against BOTH the hidden value (e.g., "12086") and the visible text
        # (e.g., "12086 - AGUR IC") to avoid re-filling an already-correct field.
        # Re-filling triggers ASP.NET autocomplete postbacks that can crash the page.
        current_project = row.get("currentProject", "").strip()
        current_project_text = row.get("currentProjectText", "").strip()
        project_code = project.split(" - ")[0].strip() if project and " - " in project else (project or "")
        project_already_set = bool(
            current_project  # hidden value is non-empty (something is selected)
            and (
                project_code in current_project  # e.g., "12086" in "12086" (exact code match)
                or project_code.lower() in current_project_text.lower()  # e.g., "agur" in "12086 - agur ic"
            )
        )
        project_needs_fill = bool(project and not project_already_set)

        if not hours_need_fill and not project_needs_fill:
            logger.info(f"  {date_text}: Already correct ({row['currentEntry']}-{row['currentExit']}), skipping")
            skipped_count += 1
            continue

        entry_id = row['entryInputId']
        exit_id = row['exitInputId']

        try:
            dismiss_modal(page, logger)

            # Fill hours if needed
            if hours_need_fill:
                fill_result = page.evaluate("""
                    (params) => {
                        const entryEl = document.getElementById(params.entryId);
                        const exitEl = document.getElementById(params.exitId);
                        if (!entryEl || !exitEl) return {ok: false, error: 'elements not found'};

                        entryEl.focus();
                        entryEl.value = params.entryTime;
                        entryEl.dispatchEvent(new Event('change', {bubbles: true}));
                        entryEl.dispatchEvent(new Event('blur', {bubbles: true}));

                        exitEl.focus();
                        exitEl.value = params.exitTime;
                        exitEl.dispatchEvent(new Event('change', {bubbles: true}));
                        exitEl.dispatchEvent(new Event('blur', {bubbles: true}));

                        return {
                            ok: true,
                            entryValue: entryEl.value,
                            exitValue: exitEl.value
                        };
                    }
                """, {"entryId": entry_id, "exitId": exit_id, "entryTime": entry_time, "exitTime": exit_time})

                logger.info(f"  {date_text}: Hours set: entry={fill_result.get('entryValue')}, exit={fill_result.get('exitValue')}")
                try:
                    page.wait_for_load_state("networkidle", timeout=10000)
                except PlaywrightTimeoutError:
                    pass
                time.sleep(1)
                dismiss_modal(page, logger)
            else:
                logger.info(f"  {date_text}: Hours already correct ({row['currentEntry']}-{row['currentExit']})")

            # Fill project for days that have hours (existing or just filled)
            project_info = ""
            has_hours_now = bool(row["currentEntry"].strip()) or hours_need_fill
            if project and project_needs_fill and has_hours_now:
                dismiss_modal(page, logger)
                result = fill_project_field(page, row, project, logger)
                project_info = f"  project={project}"
                if not result:
                    rows_needing_project_retry.append(row)

            dismiss_modal(page, logger)

            logger.info(f"  {date_text}: Filled {entry_time} - {exit_time}{project_info}")
            success_count += 1

        except Exception as e:
            logger.error(f"  {date_text}: Error filling hours: {e}")
            try:
                dismiss_modal(page, logger)
            except Exception:
                pass
            failure_count += 1
            # If the page closed, don't continue - break immediately
            if page.is_closed():
                logger.error(f"  Page closed during row {date_text}! Breaking fill loop.")
                break

    # Retry project fill for rows where autocomplete failed on first pass
    # (usually the first row, where autocomplete AJAX isn't warmed up yet)
    if rows_needing_project_retry and project:
        logger.info(f"\nRetrying project fill for {len(rows_needing_project_retry)} row(s)...")
        for row in rows_needing_project_retry:
            date_text = row["dateText"]
            dismiss_modal(page, logger)
            result = fill_project_field(page, row, project, logger)
            if result:
                logger.info(f"  {date_text}: Project retry successful")
            else:
                logger.warning(f"  {date_text}: Project retry still failed")

    # Save all changes
    if success_count > 0:
        logger.info("")
        logger.info("Saving all entries...")

        # Guard: if the page/browser is already closed, we can't save
        if page.is_closed():
            logger.error("Cannot save - page/browser is already closed!")
            logger.error(f"  {success_count} rows were filled in-memory but NOT saved.")
            return success_count, failure_count

        try:
            take_screenshot(page, "08_before_save", logger)
        except Exception:
            pass

        try:
            save_btn = page.query_selector(SEL_SAVE_BUTTON)
            if save_btn:
                save_btn.click()
                try:
                    page.wait_for_load_state("networkidle", timeout=30000)
                except PlaywrightTimeoutError:
                    pass
                time.sleep(3)
                try:
                    take_screenshot(page, "09_after_save", logger)
                except Exception:
                    pass

                # Check for any error messages after save
                post_save_check = page.evaluate("""
                    () => {
                        const errors = [];
                        // Check modal popup
                        const mpeBg = document.getElementById('MPEBehavior_backgroundElement');
                        if (mpeBg && mpeBg.style.display !== 'none') {
                            // Try to read iframe content inside the popup
                            const popup = document.getElementById('ctl00_theBasePanel');
                            if (popup) {
                                const iframe = popup.querySelector('iframe');
                                if (iframe && iframe.contentDocument) {
                                    try {
                                        errors.push('IFRAME:' + iframe.contentDocument.body.innerText);
                                    } catch(e) {
                                        errors.push('IFRAME:cross-origin');
                                    }
                                }
                                errors.push('POPUP:' + popup.innerText.substring(0, 500));
                            }
                        }
                        // Also check for any visible error text on the page
                        document.querySelectorAll('.error, .ErrorLabel, span[id*="error" i]').forEach(el => {
                            if (el.offsetParent !== null && el.innerText.trim()) {
                                errors.push('ERROR:' + el.innerText.trim());
                            }
                        });
                        return errors;
                    }
                """)
                if post_save_check:
                    logger.warning(f"Post-save messages: {post_save_check}")
                    dismiss_modal(page, logger)
                logger.info("Save completed!")
            else:
                logger.warning("Save button not found! Trying alternative selector...")
                # Try broader selector
                alt_save = page.query_selector("input[value='Save']")
                if alt_save:
                    alt_save.click()
                    page.wait_for_load_state("networkidle", timeout=30000)
                    time.sleep(3)
                    try:
                        take_screenshot(page, "09_after_save", logger)
                    except Exception:
                        pass
                    logger.info("Save completed (alt selector)!")
                else:
                    logger.error("Could not find Save button! Changes may be lost.")
                    try:
                        take_screenshot(page, "error_no_save_button", logger)
                    except Exception:
                        pass
        except Exception as save_err:
            logger.error(f"Error during save: {save_err}")
            logger.error(f"  Page may have closed. {success_count} rows were filled but save may have failed.")

    if skipped_count > 0:
        logger.info(f"  ({skipped_count} rows skipped - already filled or no inputs)")

    return success_count, failure_count


# ==============================================================================
# Main Flow
# ==============================================================================

def main():
    """Main entry point."""
    args = parse_args()
    logger = setup_logging(args.verbose)

    logger.info("=" * 60)
    logger.info("  Hilan Attendance Hours Filler")
    logger.info("  מילוי שעות נוכחות אוטומטי בחילן")
    logger.info("=" * 60)

    # Calculate workdays
    workdays = get_workdays(args.year, args.month)

    if not workdays:
        logger.warning(f"No workdays found for {args.month}/{args.year}!")
        sys.exit(0)

    # Print summary
    print_workdays_summary(workdays, args.year, args.month,
                          args.start_time, args.end_time, logger)

    # Dry run - just show the days and exit
    if args.dry_run:
        logger.info("\n[DRY RUN] No changes were made. Remove --dry-run to fill hours.")
        sys.exit(0)

    # Start browser automation
    logger.info("\nStarting browser automation...")

    with sync_playwright() as playwright:
        browser = None
        try:
            # Launch browser
            browser = playwright.chromium.launch(
                headless=args.headless,
                slow_mo=args.slow_mo,
            )
            context = browser.new_context(
                viewport={"width": 1280, "height": 900},
                locale="en-US",
            )
            page = context.new_page()

            # Set default timeouts
            page.set_default_timeout(15000)
            page.set_default_navigation_timeout(30000)

            # Step 1: Login
            login(page, args.user, args.password, logger)

            # Step 2: Navigate to attendance page
            navigate_to_attendance(page, logger)

            # Step 3: Navigate to the correct month in the calendar
            select_month_in_calendar(page, args.year, args.month, logger)

            # Step 4: Select target workdays in the calendar
            select_workdays_in_calendar(page, workdays, logger)

            # Step 5: Fill hours (and project) in the grid
            success, failure = fill_all_hours(
                page, workdays, args.start_time, args.end_time,
                args.project, args.present_weekdays, args.present_dates,
                args.vacation_days,
                args.sick_days_set, args.sick_file, logger
            )

            # Summary
            logger.info("")
            logger.info("=" * 60)
            logger.info("  SUMMARY")
            logger.info("=" * 60)
            logger.info(f"  Successfully filled: {success} days")
            logger.info(f"  Failed:              {failure} days")
            logger.info(f"  Total workdays:      {len(workdays)} days")
            logger.info("=" * 60)

            if failure > 0:
                logger.warning(f"\n{failure} days failed. Check screenshots for details.")

            take_screenshot(page, "99_final_state", logger)

            # Keep browser open for a moment to see the result
            if not args.headless:
                logger.info("\nBrowser will close in 10 seconds...")
                logger.info("(Press Ctrl+C to keep it open)")
                try:
                    time.sleep(10)
                except KeyboardInterrupt:
                    logger.info("Keeping browser open. Press Ctrl+C again to force quit.")
                    try:
                        while True:
                            time.sleep(60)
                    except KeyboardInterrupt:
                        pass

        except RuntimeError as e:
            logger.error(f"\nError: {e}")
            sys.exit(1)
        except PlaywrightTimeoutError as e:
            logger.error(f"\nTimeout error: {e}")
            logger.error("The page may be slow or the selectors may need updating.")
            if browser:
                try:
                    pages = browser.contexts[0].pages if browser.contexts else []
                    if pages:
                        take_screenshot(pages[0], "error_timeout", logger)
                except Exception:
                    pass
            sys.exit(1)
        except Exception as e:
            logger.error(f"\nUnexpected error: {e}", exc_info=True)
            if browser:
                try:
                    pages = browser.contexts[0].pages if browser.contexts else []
                    if pages:
                        take_screenshot(pages[0], "error_unexpected", logger)
                except Exception:
                    pass
            sys.exit(1)
        finally:
            if browser:
                browser.close()


if __name__ == "__main__":
    main()
