# CLAUDE.md

This project generates printable calendar PowerPoint files.

## Project Structure

- `generate_calendar.py` - Main script that creates the calendar
- `docs/` - Documentation on how the script works
- `2026_Calendar.pptx` - Generated output (not committed to git)

## Common Tasks

### Generate a calendar for a different year

Edit the `YEAR` constant in `generate_calendar.py` and run the script. The output filename is hardcoded, so also update the `output_file` variable if you want it to reflect the year.

### Change colors

Colors are defined as hex strings at the top of the script. The `set_cell_fill` function applies background colors to table cells using raw XML manipulation (python-pptx doesn't have a clean API for cell backgrounds).

### Change the week start day

The script uses `calendar.setfirstweekday(calendar.MONDAY)` to start weeks on Monday. Change this if you need a different layout, but note that `DAY_NAMES` array order must match.

## Dependencies

Only `python-pptx` is required. Install with `pip3 install python-pptx`.

## Validation

The script includes a `validate_calendar` function that checks:
1. All 12 months have tables
2. Each month has the correct number of days
3. No missing or extra days
4. First day of month is in the correct column

If validation fails, the script prints errors but still saves the file.
