# Tora Calendars

A Python script that generates printable calendar PowerPoint presentations with a Monday-Sunday week layout (weekends on the right side).

## Quick Start

```bash
# Install dependency
pip3 install python-pptx

# Generate the calendar
python3 generate_calendar.py
```

This creates `2026_Calendar.pptx` with 12 slides, one per month.

## Customization

Edit the constants at the top of `generate_calendar.py`:

- `YEAR` - Change to generate a different year
- `HEADER_BG_HEX` - Header row background color
- `WEEKEND_BG_HEX` - Saturday/Sunday column background color
- Font sizes are set in the `create_month_slide` function (title is 44pt, day headers 14pt, day numbers 18pt)

## Sample Output

A pre-generated [2026 calendar](docs/2026_Calendar.pptx) is included so you can see what the output looks like before running the script.

## Output

The generated PowerPoint has:
- 10" x 7.5" slides (good for printing)
- One month per slide
- Week runs Monday through Sunday (Saturday/Sunday on far right)
- Weekend columns have a light gray background
- Dark blue header row with day abbreviations

## Dependencies

- Python 3.x
- python-pptx

## Validation

The script automatically validates that all months have the correct number of days and that the first day of each month falls on the correct weekday column.
