#!/usr/bin/env python3
"""
Hilan Interactive Runner
========================
Interactive CLI that walks through all parameters for hilan_filler.py,
displays a visual month calendar, and launches the filler.

Usage:
    python hilan_interactive.py
"""

import calendar
import getpass
import os
import re
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

# Enable UTF-8 output on Windows
if sys.platform == "win32":
    os.system("")  # Enable ANSI escape codes on Windows 10+
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stdin.reconfigure(encoding="utf-8")
    except (AttributeError, Exception):
        pass

try:
    import holidays
except ImportError:
    print("ERROR: 'holidays' package is required. Install with:")
    print("  pip install holidays")
    sys.exit(1)

# ==============================================================================
# ANSI Color Helpers
# ==============================================================================

# Check if the terminal supports color
_COLOR_SUPPORTED = os.environ.get("NO_COLOR") is None and hasattr(sys.stdout, "isatty") and sys.stdout.isatty()


def _c(code: str, text: str) -> str:
    """Wrap text with ANSI color code if supported."""
    if not _COLOR_SUPPORTED:
        return text
    return f"\033[{code}m{text}\033[0m"


def gray(text: str) -> str:
    return _c("90", text)


def red(text: str) -> str:
    return _c("91", text)


def green(text: str) -> str:
    return _c("92", text)


def yellow(text: str) -> str:
    return _c("93", text)


def blue(text: str) -> str:
    return _c("94", text)


def magenta(text: str) -> str:
    return _c("95", text)


def cyan(text: str) -> str:
    return _c("96", text)


def bold(text: str) -> str:
    return _c("1", text)


def dim(text: str) -> str:
    return _c("2", text)


def clear_screen():
    """Clear the terminal screen."""
    os.system('cls' if sys.platform == 'win32' else 'clear')


def _visible_len(text: str) -> int:
    """Get visible length of text, ignoring ANSI escape codes."""
    return len(re.sub(r'\033\[[^m]*m', '', text))


def _rpad(text: str, width: int) -> str:
    """Right-pad text to visible width, accounting for ANSI codes."""
    return text + ' ' * max(0, width - _visible_len(text))


def _format_val(params: dict, key: str) -> str:
    """Format a parameter value for compact display."""
    current = params.get(key, "")
    if key == "password":
        return dim("****") if current else dim("—")
    elif isinstance(current, bool):
        return green("Yes") if current else "No"
    else:
        return str(current) if current else dim("—")


# ==============================================================================
# Input Helpers
# ==============================================================================

def ask(prompt: str, default: str = "", required: bool = False,
        validator=None, error_msg: str = "") -> str:
    """
    Ask the user for input with optional default, validation, and retry.
    """
    while True:
        if default:
            display = f"{prompt} [{default}]: "
        else:
            display = f"{prompt}: "

        value = input(display).strip()
        if not value:
            if default:
                return default
            if required:
                print(red("  * Required field, try again."))
                continue
            return ""

        if validator:
            ok, msg = validator(value)
            if not ok:
                print(red(f"  * {msg or error_msg or 'Invalid input, try again.'}"))
                continue

        return value


def ask_yes_no(prompt: str, default: bool = True) -> bool:
    """Ask a yes/no question."""
    suffix = " [Y/n]" if default else " [y/N]"
    while True:
        value = input(f"{prompt}{suffix}: ").strip().lower()
        if not value:
            return default
        if value in ("y", "yes"):
            return True
        if value in ("n", "no"):
            return False
        print(red("  * Enter y or n"))


# ==============================================================================
# Validators
# ==============================================================================

def validate_time(value: str):
    """Validate HH:MM time format."""
    m = re.match(r'^(\d{1,2}):(\d{2})$', value)
    if not m:
        return False, "Invalid format. Use HH:MM (e.g., 09:00)"
    h, mn = int(m.group(1)), int(m.group(2))
    if not (0 <= h <= 23 and 0 <= mn <= 59):
        return False, "Invalid time. Range: 00:00 - 23:59"
    return True, ""


def validate_month(value: str):
    """Validate month number 1-12."""
    try:
        m = int(value)
        if 1 <= m <= 12:
            return True, ""
        return False, "Month must be between 1 and 12"
    except ValueError:
        return False, "Enter a number between 1 and 12"


def validate_year(value: str):
    """Validate year."""
    try:
        y = int(value)
        if 2020 <= y <= 2030:
            return True, ""
        return False, "Year must be between 2020 and 2030"
    except ValueError:
        return False, "Enter a valid year (e.g., 2026)"


def validate_present_days(value: str):
    """Validate present-days format (comma-separated 1-5)."""
    if not value:
        return True, ""
    try:
        for part in value.split(","):
            d = int(part.strip())
            if d not in (1, 2, 3, 4, 5):
                return False, f"Day {d} is invalid. Use: 1=Sun, 2=Mon, 3=Tue, 4=Wed, 5=Thu"
        return True, ""
    except ValueError:
        return False, "Invalid format. Enter comma-separated numbers (e.g., 2,4)"


def parse_day_ranges(value: str) -> set:
    """Parse day ranges like '1-3,15,20-22' into a set of day numbers."""
    days = set()
    for part in value.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start, end = part.split("-", 1)
            for d in range(int(start.strip()), int(end.strip()) + 1):
                days.add(d)
        else:
            days.add(int(part))
    return days


def validate_day_ranges(value: str):
    """Validate day range format (e.g., '1-3,15,20-22')."""
    if not value:
        return True, ""
    try:
        days = parse_day_ranges(value)
        for d in days:
            if d < 1 or d > 31:
                return False, f"Day {d} is invalid (1-31)"
        return True, ""
    except ValueError:
        return False, "Invalid format. Examples: 8-12 or 1,3,15 or 1-3,15,20-22"


# ==============================================================================
# Calendar Display
# ==============================================================================

# Week order for Israeli calendar: Sun, Mon, Tue, Wed, Thu, Fri, Sat
# Python weekday(): Mon=0, Tue=1, Wed=2, Thu=3, Fri=4, Sat=5, Sun=6
WEEK_ORDER = [6, 0, 1, 2, 3, 4, 5]  # Sun, Mon, Tue, Wed, Thu, Fri, Sat

DAY_NAMES_SHORT = {
    6: "Sun", 0: "Mon", 1: "Tue", 2: "Wed", 3: "Thu", 4: "Fri", 5: "Sat"
}


def get_israeli_holidays(year: int) -> dict:
    """Get Israeli holidays for the given year as {date: name} in English."""
    il_holidays = holidays.Israel(years=year, language="en_US")
    return dict(il_holidays)


def display_calendar(year: int, month: int, vacation_days: set = None,
                     sick_days: set = None, present_weekdays: set = None,
                     present_dates: set = None, rd_days: set = None):
    """
    Display a visual calendar for the given month.
    Colors:
      - Green: regular workday (w.home)
      - Cyan: presence day (in office) - weekly recurring or specific date
      - Magenta: vacation
      - Yellow: sick
      - Blue: reserve duty (מילואים)
      - Red: holiday
      - Gray: weekend (Fri/Sat)
    """
    vacation_days = vacation_days or set()
    sick_days = sick_days or set()
    present_weekdays = present_weekdays or set()
    present_dates = present_dates or set()
    rd_days = rd_days or set()

    holiday_map = get_israeli_holidays(year)
    _, num_days = calendar.monthrange(year, month)

    month_name_en = calendar.month_name[month]

    print()
    print(bold(f"  === {month_name_en} {year} ==="))
    print()

    # Header: day names
    header = "  "
    for wd in WEEK_ORDER:
        day_name = DAY_NAMES_SHORT[wd]
        if wd in (4, 5):  # Fri, Sat
            header += gray(f" {day_name:>4} ")
        else:
            header += f" {day_name:>4} "
    print(header)
    print("  " + "-" * 42)

    # Build rows
    first_day = date(year, month, 1)
    # Find position of day 1 in our week order
    first_weekday = first_day.weekday()  # Python weekday
    start_pos = WEEK_ORDER.index(first_weekday)

    row = "  "
    # Pad empty cells before day 1
    for _ in range(start_pos):
        row += "      "

    col = start_pos
    for day_num in range(1, num_days + 1):
        d = date(year, month, day_num)
        wd = d.weekday()

        cell = f"{day_num:>4}"

        is_holiday = d in holiday_map
        is_weekend = wd in (4, 5)  # Fri, Sat
        is_vacation = day_num in vacation_days
        is_sick = day_num in sick_days
        is_rd = day_num in rd_days
        is_workday = wd in (6, 0, 1, 2, 3) and not is_holiday  # Sun-Thu, no holiday
        is_presence = is_workday and (wd in present_weekdays or day_num in present_dates)

        # Determine display style (priority: vacation > sick > rd > holiday > presence > workday > weekend)
        if is_vacation and is_workday:
            cell = magenta(cell)
        elif is_sick and is_workday:
            cell = yellow(cell)
        elif is_rd and is_workday:
            cell = blue(cell)
        elif is_holiday:
            cell = red(cell)
        elif is_weekend:
            cell = gray(cell)
        elif is_presence:
            cell = cyan(cell)
        elif is_workday:
            cell = green(cell)

        row += f" {cell} "
        col += 1

        if col >= 7:
            print(row)
            row = "  "
            col = 0

    if col > 0:
        print(row)

    # Legend
    print()
    print("  " + bold("Legend:"))
    legend_parts = [
        green("##") + " Workday (w.home)",
    ]
    if present_weekdays or present_dates:
        legend_parts.append(cyan("##") + " Presence (in office)")
    if vacation_days:
        legend_parts.append(magenta("##") + " Vacation")
    if sick_days:
        legend_parts.append(yellow("##") + " Sick")
    if rd_days:
        legend_parts.append(blue("##") + " Reserve Duty")
    legend_parts.append(red("##") + " Holiday")
    legend_parts.append(gray("##") + " Weekend (Fri/Sat)")

    for part in legend_parts:
        print(f"    {part}")

    # Show holiday names
    month_holidays = {d: name for d, name in holiday_map.items()
                      if d.month == month and d.year == year}
    if month_holidays:
        print()
        print("  " + bold("Holidays this month:"))
        for d in sorted(month_holidays):
            print(f"    {d.day:>2}/{d.month:>02} - {red(month_holidays[d])}")

    print()


def display_compact_calendar(year: int, month: int, vacation_days: set = None,
                             sick_days: set = None, present_weekdays: set = None,
                             present_dates: set = None, rd_days: set = None):
    """Display a compact calendar for the edit screen (fewer lines, inline legend)."""
    vacation_days = vacation_days or set()
    sick_days = sick_days or set()
    present_weekdays = present_weekdays or set()
    present_dates = present_dates or set()
    rd_days = rd_days or set()

    holiday_map = get_israeli_holidays(year)
    _, num_days = calendar.monthrange(year, month)
    month_name_en = calendar.month_name[month]

    print(bold(f"  {month_name_en} {year}"))

    # Compact header (3-char cells)
    header = " "
    for wd in WEEK_ORDER:
        name = DAY_NAMES_SHORT[wd][:2]
        if wd in (4, 5):
            header += gray(f"{name:>3}")
        else:
            header += f"{name:>3}"
    print(header)

    # Days (3-char cells)
    first_day = date(year, month, 1)
    first_weekday = first_day.weekday()
    start_pos = WEEK_ORDER.index(first_weekday)

    row = " "
    for _ in range(start_pos):
        row += "   "

    col = start_pos
    for day_num in range(1, num_days + 1):
        d = date(year, month, day_num)
        wd = d.weekday()
        cell = f"{day_num:>3}"

        is_holiday = d in holiday_map
        is_weekend = wd in (4, 5)
        is_vacation = day_num in vacation_days
        is_sick = day_num in sick_days
        is_rd = day_num in rd_days
        is_workday = wd in (6, 0, 1, 2, 3) and not is_holiday
        is_presence = is_workday and (wd in present_weekdays or day_num in present_dates)

        if is_vacation and is_workday:
            cell = magenta(cell)
        elif is_sick and is_workday:
            cell = yellow(cell)
        elif is_rd and is_workday:
            cell = blue(cell)
        elif is_holiday:
            cell = red(cell)
        elif is_weekend:
            cell = gray(cell)
        elif is_presence:
            cell = cyan(cell)
        elif is_workday:
            cell = green(cell)

        row += cell
        col += 1
        if col >= 7:
            print(row)
            row = " "
            col = 0

    if col > 0:
        print(row)

    # Compact single-line legend
    parts = [green("##") + " Work"]
    if present_weekdays or present_dates:
        parts.append(cyan("##") + " Office")
    if vacation_days:
        parts.append(magenta("##") + " Vacation")
    if sick_days:
        parts.append(yellow("##") + " Sick")
    if rd_days:
        parts.append(blue("##") + " RD")
    parts.append(red("##") + " Holiday")
    parts.append(gray("##") + " Wknd")
    print("  " + "  ".join(parts))

    # Compact holiday names (single line)
    month_holidays = {d: name for d, name in holiday_map.items()
                      if d.month == month and d.year == year}
    if month_holidays:
        hol_parts = [f"{d.day}/{d.month:02} {red(month_holidays[d])}" for d in sorted(month_holidays)]
        print("  " + ", ".join(hol_parts))


# ==============================================================================
# Summary Display
# ==============================================================================

def display_summary(params: dict):
    """Display a summary of all collected parameters."""
    print()
    print(bold("=" * 60))
    print(bold("  Summary"))
    print(bold("=" * 60))
    print()
    print(f"  {'Employee #:':<20} {params['user']}")
    print(f"  {'Password:':<20} {'*' * len(params['password'])}")
    print(f"  {'Month:':<20} {calendar.month_name[params['month']]} {params['year']} ({params['month']}/{params['year']})")
    print(f"  {'Project:':<20} {params['project']}")
    print(f"  {'Entry time:':<20} {params['start_time']}")
    print(f"  {'Exit time:':<20} {params['end_time']}")

    if params.get('present_days'):
        day_map = {"1": "Sun", "2": "Mon", "3": "Tue", "4": "Wed", "5": "Thu"}
        day_names = [day_map.get(d.strip(), d.strip()) for d in params['present_days'].split(",")]
        print(f"  {'Office days (weekly):':<20} {', '.join(day_names)} ({params['present_days']})")
    else:
        print(f"  {'Office days (weekly):':<20} {dim('None')}")

    if params.get('present_dates'):
        print(f"  {'Office dates:':<20} {params['present_dates']}")
    else:
        print(f"  {'Office dates:':<20} {dim('None')}")

    if not params.get('present_days') and not params.get('present_dates'):
        print(f"  {'':>20} {dim('(all workdays will be w.home)')}")

    if params.get('vacation'):
        print(f"  {'Vacation days:':<20} {params['vacation']}")
    else:
        print(f"  {'Vacation days:':<20} {dim('None')}")

    if params.get('sick_days'):
        sick_count = len(parse_day_ranges(params['sick_days']))
        sick_type = "sick day declaration" if sick_count <= 2 else "sick (with certificate)"
        print(f"  {'Sick days:':<20} {params['sick_days']} ({sick_type})")
        if params.get('sick_file'):
            print(f"  {'Sick file:':<20} {params['sick_file']}")
    else:
        print(f"  {'Sick days:':<20} {dim('None')}")

    if params.get('rd_days'):
        print(f"  {'RD days (מילואים):':<20} {params['rd_days']}")
        if params.get('rd_file'):
            print(f"  {'RD file:':<20} {params['rd_file']}")
    else:
        print(f"  {'RD days (מילואים):':<20} {dim('None')}")

    print(f"  {'Headless:':<20} {'Yes' if params.get('headless') else 'No (browser visible)'}")
    print(f"  {'Dry-run:':<20} {'Yes (no changes)' if params.get('dry_run') else 'No (real fill)'}")
    print()

    # Show annotated calendar
    # Convert present_days string to present_weekdays set (Python weekdays)
    user_to_python_weekday = {1: 6, 2: 0, 3: 1, 4: 2, 5: 3}
    present_weekdays = set()
    if params.get('present_days'):
        for d in params['present_days'].split(","):
            d = d.strip()
            if d:
                present_weekdays.add(user_to_python_weekday[int(d)])

    vacation_set = parse_day_ranges(params['vacation']) if params.get('vacation') else set()
    sick_set = parse_day_ranges(params['sick_days']) if params.get('sick_days') else set()
    present_dates_set = parse_day_ranges(params['present_dates']) if params.get('present_dates') else set()
    rd_set = parse_day_ranges(params['rd_days']) if params.get('rd_days') else set()

    display_calendar(params['year'], params['month'],
                     vacation_days=vacation_set,
                     sick_days=sick_set,
                     present_weekdays=present_weekdays,
                     present_dates=present_dates_set,
                     rd_days=rd_set)


def build_command(params: dict) -> list[str]:
    """Build the command-line arguments list for hilan_filler.py."""
    script_dir = Path(__file__).parent
    script_path = str(script_dir / "hilan_filler.py")

    cmd = [sys.executable, script_path]
    cmd += ["-u", params["user"]]
    cmd += ["-p", params["password"]]
    cmd += ["--project", params["project"]]
    cmd += ["--month", str(params["month"])]
    cmd += ["--year", str(params["year"])]
    cmd += ["--start-time", params["start_time"]]
    cmd += ["--end-time", params["end_time"]]

    if params.get("present_days"):
        cmd += ["--present-days", params["present_days"]]
    if params.get("present_dates"):
        cmd += ["--present-dates", params["present_dates"]]
    if params.get("vacation"):
        cmd += ["--vacation", params["vacation"]]
    if params.get("sick_days"):
        cmd += ["--sick-days", params["sick_days"]]
    if params.get("sick_file"):
        cmd += ["--sick-file", params["sick_file"]]
    if params.get("rd_days"):
        cmd += ["--rd-days", params["rd_days"]]
    if params.get("rd_file"):
        cmd += ["--rd-file", params["rd_file"]]
    if params.get("headless"):
        cmd += ["--headless"]
    if params.get("dry_run"):
        cmd += ["--dry-run"]

    return cmd


def display_command(params: dict):
    """Display the command that will be run (with masked password)."""
    cmd = build_command(params)
    # Mask the password in the display
    display_parts = []
    mask_next = False
    for idx, part in enumerate(cmd):
        if mask_next:
            display_parts.append("****")
            mask_next = False
        elif part == "-p":
            display_parts.append(part)
            mask_next = True
        elif idx == 0:
            # Show 'python' instead of the full interpreter path
            display_parts.append("python")
        else:
            # Quote parts with spaces
            if " " in part:
                display_parts.append(f'"{part}"')
            else:
                display_parts.append(part)

    print(dim("  Command: " + " ".join(display_parts)))
    print()


# ==============================================================================
# Main Interactive Flow
# ==============================================================================

def main():
    """Main interactive flow."""

    while True:
        params = collect_params()
        if params is None:
            print(yellow("\n  Cancelled."))
            sys.exit(0)

        # Show summary
        clear_screen()
        display_summary(params)
        display_command(params)

        # Confirm
        choice = ask(
            bold("What to do?") + " (r=Run, e=Edit params, q=Quit)",
            default="r",
        ).lower()

        if choice in ("q", "quit"):
            print(yellow("\n  Exiting."))
            sys.exit(0)
        elif choice in ("e", "edit"):
            params = edit_params(params)
            continue
        elif choice in ("r", "run"):
            break
        else:
            break

    # Run hilan_filler.py
    print()
    print(bold("=" * 60))
    print(bold("  Running hilan_filler.py..."))
    print(bold("=" * 60))
    print()

    cmd = build_command(params)
    result = subprocess.run(cmd)
    sys.exit(result.returncode)


def edit_params(params: dict) -> dict:
    """Let the user pick which parameter to edit without restarting."""
    EDIT_FIELDS = [
        ("1",  "Employee#",  "user"),
        ("2",  "Password",   "password"),
        ("3",  "Month",      "month"),
        ("4",  "Year",       "year"),
        ("5",  "Project",    "project"),
        ("6",  "Entry",      "start_time"),
        ("7",  "Exit",       "end_time"),
        ("8",  "Office wk",  "present_days"),
        ("9",  "Office dt",  "present_dates"),
        ("10", "Vacation",   "vacation"),
        ("11", "Sick days",  "sick_days"),
        ("12", "RD days",    "rd_days"),
        ("13", "Headless",   "headless"),
        ("14", "Dry-run",    "dry_run"),
    ]

    while True:
        # Clear screen before showing the edit table
        clear_screen()

        # Compute calendar display sets from current params
        _u2pw = {1: 6, 2: 0, 3: 1, 4: 2, 5: 3}
        pw_set = set()
        if params.get('present_days'):
            for d in params['present_days'].split(","):
                d = d.strip()
                if d:
                    pw_set.add(_u2pw[int(d)])
        vac_set = parse_day_ranges(params['vacation']) if params.get('vacation') else set()
        sick_set = parse_day_ranges(params['sick_days']) if params.get('sick_days') else set()
        pd_set = parse_day_ranges(params['present_dates']) if params.get('present_dates') else set()
        rd_set = parse_day_ranges(params['rd_days']) if params.get('rd_days') else set()

        print()
        display_compact_calendar(params["year"], params["month"],
                                 vacation_days=vac_set, sick_days=sick_set,
                                 present_weekdays=pw_set, present_dates=pd_set,
                                 rd_days=rd_set)

        # Two-column parameter table
        print()
        left = EDIT_FIELDS[:7]
        right = EDIT_FIELDS[7:]
        for i in range(len(left)):
            num_l, lbl_l, key_l = left[i]
            left_cell = f"  {num_l:>2}) " + _rpad(lbl_l, 10) + _format_val(params, key_l)

            if i < len(right):
                num_r, lbl_r, key_r = right[i]
                right_cell = f"{num_r:>2}) " + _rpad(lbl_r, 10) + _format_val(params, key_r)
                print(_rpad(left_cell, 30) + right_cell)
            else:
                print(left_cell)
        print(f"   d) {bold('Done')}")
        print()

        choice = ask("Edit#", default="d")

        if choice.lower() == "d":
            # Validate required fields
            if not params.get("project"):
                print(red("  * Project is required. Please set it (option 5) before continuing."))
                continue
            # Validate no overlap between vacation and sick
            if params.get("vacation") and params.get("sick_days"):
                vac_set = parse_day_ranges(params["vacation"])
                sick_set = parse_day_ranges(params["sick_days"])
                overlap = vac_set & sick_set
                if overlap:
                    print(red(f"  * Vacation and sick days overlap on: {sorted(overlap)}"))
                    print(red("  * Fix vacation (10) or sick days (11) before continuing."))
                    continue
            # Validate no overlap between rd and vacation/sick
            if params.get("rd_days"):
                rd_set = parse_day_ranges(params["rd_days"])
                if params.get("vacation"):
                    vac_set = parse_day_ranges(params["vacation"])
                    overlap = rd_set & vac_set
                    if overlap:
                        print(red(f"  * RD and vacation overlap on: {sorted(overlap)}"))
                        print(red("  * Fix RD days (12) or vacation (10) before continuing."))
                        continue
                if params.get("sick_days"):
                    sick_set = parse_day_ranges(params["sick_days"])
                    overlap = rd_set & sick_set
                    if overlap:
                        print(red(f"  * RD and sick days overlap on: {sorted(overlap)}"))
                        print(red("  * Fix RD days (12) or sick days (11) before continuing."))
                        continue
            # Validate no overlap between present_dates and vacation/sick/rd
            if params.get("present_dates"):
                pres_set = parse_day_ranges(params["present_dates"])
                if params.get("vacation"):
                    vac_set = parse_day_ranges(params["vacation"])
                    overlap = pres_set & vac_set
                    if overlap:
                        print(red(f"  * Present dates and vacation overlap on: {sorted(overlap)}"))
                        print(red("  * Fix present dates (9) or vacation (10) before continuing."))
                        continue
                if params.get("sick_days"):
                    sick_set = parse_day_ranges(params["sick_days"])
                    overlap = pres_set & sick_set
                    if overlap:
                        print(red(f"  * Present dates and sick days overlap on: {sorted(overlap)}"))
                        print(red("  * Fix present dates (9) or sick days (11) before continuing."))
                        continue
                if params.get("rd_days"):
                    rd_set = parse_day_ranges(params["rd_days"])
                    overlap = pres_set & rd_set
                    if overlap:
                        print(red(f"  * Present dates and RD days overlap on: {sorted(overlap)}"))
                        print(red("  * Fix present dates (9) or RD days (12) before continuing."))
                        continue
            break

        # Find the selected field
        field = None
        for num, label, key in EDIT_FIELDS:
            if choice == num:
                field = (num, label, key)
                break

        if not field:
            print(red(f"  * Invalid choice: {choice}"))
            continue

        _, label, key = field
        print()

        if key == "user":
            params["user"] = ask("Employee number", default=params["user"], required=True)

        elif key == "password":
            try:
                new_pw = getpass.getpass("New password (Enter to keep current): ")
                if new_pw:
                    confirm_pw = getpass.getpass("Confirm password: ")
                    if new_pw == confirm_pw:
                        params["password"] = new_pw
                    else:
                        print(red("  * Passwords do not match. Password not changed."))
                        continue
                else:
                    print(dim("  (keeping current password)"))
            except (EOFError, KeyboardInterrupt):
                print(dim("  (keeping current password)"))

        elif key == "month":
            val = ask("Month (1-12)", default=str(params["month"]), validator=validate_month)
            params["month"] = int(val)

        elif key == "year":
            val = ask("Year", default=str(params["year"]), validator=validate_year)
            params["year"] = int(val)

        elif key == "project":
            params["project"] = ask("Project", default=params.get("project", ""), required=True)

        elif key == "start_time":
            params["start_time"] = ask("Entry time (HH:MM)", default=params["start_time"], validator=validate_time)

        elif key == "end_time":
            params["end_time"] = ask("Exit time (HH:MM)", default=params["end_time"], validator=validate_time)

        elif key == "present_days":
            print(dim("  Format: 1=Sun, 2=Mon, 3=Tue, 4=Wed, 5=Thu"))
            print(dim("  Example: 2,4  (Mon and Wed in office every week)"))
            print(dim("  Leave empty if no fixed weekly days"))
            params["present_days"] = ask(
                "Weekly office days",
                default=params.get("present_days", ""),
                validator=validate_present_days,
            )

        elif key == "present_dates":
            print(dim("  Specific days in the month you were in the office."))
            print(dim("  Supports ranges: 5,12,19 or 10-14 or 5,10-14,19"))
            print(dim("  Leave empty to clear"))
            params["present_dates"] = ask(
                "Specific office dates",
                default=params.get("present_dates", ""),
                validator=validate_day_ranges,
            )
            # Check overlap with vacation
            if params.get("vacation") and params.get("present_dates"):
                pres_set = parse_day_ranges(params["present_dates"])
                vac_set = parse_day_ranges(params["vacation"])
                overlap = pres_set & vac_set
                if overlap:
                    print(red(f"  * Warning: overlaps with vacation on days: {sorted(overlap)}"))
                    print(red("  * Please fix present dates or vacation days."))
            # Check overlap with sick
            if params.get("sick_days") and params.get("present_dates"):
                pres_set = parse_day_ranges(params["present_dates"])
                sick_set = parse_day_ranges(params["sick_days"])
                overlap = pres_set & sick_set
                if overlap:
                    print(red(f"  * Warning: overlaps with sick days on days: {sorted(overlap)}"))
                    print(red("  * Please fix present dates or sick days."))
            # Check overlap with RD
            if params.get("rd_days") and params.get("present_dates"):
                pres_set = parse_day_ranges(params["present_dates"])
                rd_set = parse_day_ranges(params["rd_days"])
                overlap = pres_set & rd_set
                if overlap:
                    print(red(f"  * Warning: overlaps with RD days on days: {sorted(overlap)}"))
                    print(red("  * Please fix present dates or RD days."))

        elif key == "vacation":
            print(dim("  Supports ranges: 8-12 or 1,3,15 or 1-3,15,20-22"))
            print(dim("  Leave empty to clear"))
            params["vacation"] = ask(
                "Vacation days",
                default=params.get("vacation", ""),
                validator=validate_day_ranges,
            )

        elif key == "sick_days":
            print(dim("  1-2 days = sick day declaration, 3+ = sick (needs certificate)"))
            print(dim("  Supports ranges: 8-10 or 5,6 or 1-3,15"))
            print(dim("  Leave empty to clear"))
            params["sick_days"] = ask(
                "Sick days",
                default=params.get("sick_days", ""),
                validator=validate_day_ranges,
            )
            # Check overlap with vacation
            if params["vacation"] and params["sick_days"]:
                vac_set = parse_day_ranges(params["vacation"])
                sick_set = parse_day_ranges(params["sick_days"])
                overlap = vac_set & sick_set
                if overlap:
                    print(red(f"  * Warning: overlaps with vacation on days: {sorted(overlap)}"))
                    print(red("  * Please fix vacation or sick days."))
            # Handle sick file
            params["sick_file"] = ""
            if params["sick_days"]:
                sick_count = len(parse_day_ranges(params["sick_days"]))
                if sick_count > 2:
                    print(yellow(f"  * {sick_count} sick days - certificate required."))
                    while True:
                        sick_file = ask("Path to sick certificate", required=True)
                        if Path(sick_file).exists():
                            params["sick_file"] = sick_file
                            break
                        else:
                            print(red(f"  * File not found: {sick_file}"))

        elif key == "rd_days":
            print(dim("  Reserve duty days (מילואים). Requires a file (צו מילואים)."))
            print(dim("  Supports ranges: 5-7 or 10,11,12 or 5-7,15"))
            print(dim("  Leave empty to clear"))
            params["rd_days"] = ask(
                "Reserve duty days",
                default=params.get("rd_days", ""),
                validator=validate_day_ranges,
            )
            # Check overlap with vacation
            if params.get("rd_days") and params.get("vacation"):
                rd_set = parse_day_ranges(params["rd_days"])
                vac_set = parse_day_ranges(params["vacation"])
                overlap = rd_set & vac_set
                if overlap:
                    print(red(f"  * Warning: overlaps with vacation on days: {sorted(overlap)}"))
                    print(red("  * Please fix RD days or vacation days."))
            # Check overlap with sick
            if params.get("rd_days") and params.get("sick_days"):
                rd_set = parse_day_ranges(params["rd_days"])
                sick_set = parse_day_ranges(params["sick_days"])
                overlap = rd_set & sick_set
                if overlap:
                    print(red(f"  * Warning: overlaps with sick days on days: {sorted(overlap)}"))
                    print(red("  * Please fix RD days or sick days."))
            # Handle RD file
            params["rd_file"] = ""
            if params["rd_days"]:
                rd_count = len(parse_day_ranges(params["rd_days"]))
                print(yellow(f"  * {rd_count} reserve duty day(s) - file required."))
                while True:
                    rd_file = ask("Path to reserve duty order file", required=True)
                    if Path(rd_file).exists():
                        params["rd_file"] = rd_file
                        break
                    else:
                        print(red(f"  * File not found: {rd_file}"))

        elif key == "headless":
            params["headless"] = ask_yes_no("Run headless (no browser window)?",
                                            default=params.get("headless", False))

        elif key == "dry_run":
            params["dry_run"] = ask_yes_no("Dry-run only (no changes)?",
                                           default=params.get("dry_run", False))

        print(green(f"  Updated {label}."))

    return params


def _print_header():
    """Print the app header."""
    print(bold("=" * 60))
    print(bold("  Hilan Interactive Runner"))
    print(bold("  Automatic hours filler for Hilan"))
    print(bold("=" * 60))
    print()


def collect_params() -> dict | None:
    """Collect all parameters interactively. Returns dict or None if cancelled."""
    params = {}
    today = date.today()

    # --- Step 1: Credentials ---
    clear_screen()
    _print_header()
    print(bold("  Step 1/8 — Login Credentials"))
    print()

    params["user"] = ask("Employee number", required=True)

    clear_screen()
    _print_header()
    print(bold("  Step 1/8 — Login Credentials"))
    print()
    print(f"  Employee#: {params['user']}")
    print()

    try:
        while True:
            pw1 = getpass.getpass("Password: ")
            if not pw1:
                print(red("  * Password is required."))
                continue
            pw2 = getpass.getpass("Confirm password: ")
            if pw1 != pw2:
                print(red("  * Passwords do not match. Try again."))
                continue
            params["password"] = pw1
            break
    except (EOFError, KeyboardInterrupt):
        return None

    # --- Step 2: Month ---
    clear_screen()
    _print_header()
    print(bold("  Step 2/8 — Select Month"))
    print()

    params["month"] = int(ask(
        "Month (1-12)",
        default=str(today.month),
        validator=validate_month,
    ))

    clear_screen()
    _print_header()
    print(bold("  Step 2/8 — Select Month"))
    print()
    print(f"  Month: {params['month']}")
    print()

    params["year"] = int(ask(
        "Year",
        default=str(today.year),
        validator=validate_year,
    ))

    # --- Step 3: Set defaults ---
    params["start_time"] = "09:00"
    params["end_time"] = "18:00"
    params["present_dates"] = ""
    params["sick_file"] = ""
    params["headless"] = False
    params["dry_run"] = False

    # --- Step 4: Project ---
    clear_screen()
    _print_header()
    print(bold(f"  Step 3/8 — Project  ({calendar.month_name[params['month']]} {params['year']})"))
    print()
    params["project"] = ask("Project code (e.g., 12086 - AGUR IC)", required=True)

    # --- Step 5: Vacation days ---
    clear_screen()
    _print_header()
    print(bold(f"  Step 4/8 — Vacation  ({calendar.month_name[params['month']]} {params['year']})"))
    print()
    print(dim("  Supports ranges: 8-12 or 1,3,15 or 1-3,15,20-22"))
    print(dim("  Leave empty if no vacation this month"))
    print()
    params["vacation"] = ask(
        "Vacation days",
        default="",
        validator=validate_day_ranges,
    )

    # --- Step 6: Sick days ---
    clear_screen()
    _print_header()
    print(bold(f"  Step 5/8 — Sick Days  ({calendar.month_name[params['month']]} {params['year']})"))
    print()
    print(dim("  1-2 days = sick day declaration, 3+ = sick (needs certificate)"))
    print(dim("  Supports ranges: 8-10 or 5,6 or 1-3,15"))
    print(dim("  Leave empty if no sick days this month"))
    print()
    params["sick_days"] = ask(
        "Sick days",
        default="",
        validator=validate_day_ranges,
    )
    # Handle sick file if needed
    if params["sick_days"]:
        sick_count = len(parse_day_ranges(params["sick_days"]))
        if sick_count > 2:
            print(yellow(f"  * {sick_count} sick days - certificate required."))
            while True:
                sick_file = ask("Path to sick certificate", required=True)
                if Path(sick_file).exists():
                    params["sick_file"] = sick_file
                    break
                else:
                    print(red(f"  * File not found: {sick_file}"))
    # Check overlap with vacation
    if params.get("vacation") and params.get("sick_days"):
        vac_set = parse_day_ranges(params["vacation"])
        sick_set = parse_day_ranges(params["sick_days"])
        overlap = vac_set & sick_set
        if overlap:
            print(red(f"  * Warning: vacation and sick overlap on days: {sorted(overlap)}"))
            print(red("  * You can fix this in the edit screen."))
            input(dim("  Press Enter to continue..."))

    # --- Step 6.5: Reserve duty days ---
    clear_screen()
    _print_header()
    print(bold(f"  Step 6/8 — Reserve Duty (מילואים)  ({calendar.month_name[params['month']]} {params['year']})"))
    print()
    print(dim("  Supports ranges: 5-7 or 10,11,12 or 5-7,15"))
    print(dim("  Requires attaching a reserve duty order file (צו מילואים)"))
    print(dim("  Leave empty if no reserve duty this month"))
    print()
    params["rd_days"] = ask(
        "Reserve duty days",
        default="",
        validator=validate_day_ranges,
    )
    params["rd_file"] = ""
    if params["rd_days"]:
        rd_count = len(parse_day_ranges(params["rd_days"]))
        print(yellow(f"  * {rd_count} reserve duty day(s) - file required."))
        while True:
            rd_file = ask("Path to reserve duty order file", required=True)
            if Path(rd_file).exists():
                params["rd_file"] = rd_file
                break
            else:
                print(red(f"  * File not found: {rd_file}"))
    # Check overlap with vacation/sick
    if params.get("rd_days"):
        rd_set = parse_day_ranges(params["rd_days"])
        if params.get("vacation"):
            vac_set = parse_day_ranges(params["vacation"])
            overlap = rd_set & vac_set
            if overlap:
                print(red(f"  * Warning: RD and vacation overlap on days: {sorted(overlap)}"))
                print(red("  * You can fix this in the edit screen."))
                input(dim("  Press Enter to continue..."))
        if params.get("sick_days"):
            sick_set = parse_day_ranges(params["sick_days"])
            overlap = rd_set & sick_set
            if overlap:
                print(red(f"  * Warning: RD and sick overlap on days: {sorted(overlap)}"))
                print(red("  * You can fix this in the edit screen."))
                input(dim("  Press Enter to continue..."))

    # --- Step 7: Office days ---
    clear_screen()
    _print_header()
    print(bold(f"  Step 7/8 — Office Days  ({calendar.month_name[params['month']]} {params['year']})"))
    print()
    print(dim("  Format: 1=Sun, 2=Mon, 3=Tue, 4=Wed, 5=Thu"))
    print(dim("  Example: 2,4  (Mon and Wed in office every week)"))
    print(dim("  Leave empty if working from home all days"))
    print()
    params["present_days"] = ask(
        "Weekly office days",
        default="2,4",
        validator=validate_present_days,
    )

    # --- Display the calendar ---
    clear_screen()
    _print_header()
    print(bold(f"  Step 8/8 — Review  ({calendar.month_name[params['month']]} {params['year']})"))

    user_to_python_weekday = {1: 6, 2: 0, 3: 1, 4: 2, 5: 3}
    pw_set = set()
    if params.get('present_days'):
        for d in params['present_days'].split(","):
            d = d.strip()
            if d:
                pw_set.add(user_to_python_weekday[int(d)])
    vac_set = parse_day_ranges(params['vacation']) if params.get('vacation') else set()
    sick_set = parse_day_ranges(params['sick_days']) if params.get('sick_days') else set()
    rd_set = parse_day_ranges(params['rd_days']) if params.get('rd_days') else set()

    display_calendar(params['year'], params['month'],
                     vacation_days=vac_set, sick_days=sick_set,
                     present_weekdays=pw_set, rd_days=rd_set)

    # --- Go to edit screen for fine-tuning ---
    params = edit_params(params)

    return params


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(yellow("\n\n  Cancelled (Ctrl+C)."))
        sys.exit(0)
