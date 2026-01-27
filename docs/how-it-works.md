# How the Calendar Generator Works

This document explains the implementation of `generate_calendar.py` so you can modify it or create similar calendars in the future.

## Overview

The script uses Python's built-in `calendar` module to get the correct days for each month, then uses `python-pptx` to create a PowerPoint presentation with one slide per month. Each slide contains a title (month and year) and a table showing the calendar grid.

## Key Design Decisions

### Monday-Sunday Week Layout

Most American calendars start on Sunday, but this one starts on Monday so that Saturday and Sunday appear together on the right side. This is controlled by:

```python
calendar.setfirstweekday(calendar.MONDAY)
```

The `calendar.monthcalendar(year, month)` function then returns weeks as lists where index 0 is Monday and index 6 is Sunday.

### Table-Based Layout

Each month is rendered as a table with:
- 7 columns (one per day of week)
- Variable rows depending on how many weeks the month spans (typically 5-6 rows plus 1 header row)

Tables in PowerPoint are easier to style consistently than individual text boxes and align naturally to a grid.

### Cell Background Colors

The `python-pptx` library doesn't have a direct API for setting table cell background colors. The script works around this by manipulating the underlying XML:

```python
def set_cell_fill(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for child in list(tcPr):
        if 'solidFill' in child.tag:
            tcPr.remove(child)
    solidFill = parse_xml(
        f'<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:srgbClr val="{hex_color}"/></a:solidFill>'
    )
    tcPr.append(solidFill)
```

This accesses the cell's internal XML element (`_tc`), gets or creates its properties element (`tcPr`), removes any existing fill, and adds a new solid fill with the specified color.

## Script Structure

### Constants

```python
YEAR = 2026
MONTH_NAMES = ["January", "February", ...]
DAY_NAMES = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
HEADER_BG_HEX = "2C3E50"   # Dark blue
WEEKEND_BG_HEX = "ECF0F1"  # Light gray
```

Change `YEAR` to generate a different year's calendar. The color values are hex RGB without the `#` prefix.

### create_month_slide(prs, month)

This function creates a single slide for one month:

1. Adds a blank slide to the presentation
2. Creates a text box with the month/year title
3. Gets the calendar grid using `calendar.monthcalendar(YEAR, month)`
4. Creates a table with the appropriate number of rows
5. Fills the header row with day names (Mon-Sun)
6. Fills each cell with the day number (or empty for days outside the month)
7. Applies background colors (gray for weekend columns, white for weekday columns)

### validate_calendar(prs)

After generating all slides, this function checks each month:

1. Finds the table on each slide
2. Extracts all day numbers from the table cells
3. Verifies the count matches the expected number of days in that month
4. Checks that no days are missing or duplicated
5. Confirms the first day of the month appears in the correct column

This catches bugs like off-by-one errors or incorrect weekday calculations.

### main()

Orchestrates the whole process:
1. Creates a new presentation with custom slide dimensions (10" x 7.5")
2. Calls `create_month_slide` for months 1-12
3. Runs validation
4. Saves the PowerPoint file

## Modifying for Future Years

To generate a 2027 calendar:

1. Change `YEAR = 2027` at the top of the script
2. Optionally change `output_file = "2027_Calendar.pptx"` near the bottom
3. Run the script

No other changes needed. Python's `calendar` module handles leap years and weekday calculations automatically.

## Styling Changes

### Colors

Edit the hex color constants. Some options:
- Header: `"1A5276"` (darker blue), `"922B21"` (burgundy), `"1E8449"` (green)
- Weekend: `"FADBD8"` (light pink), `"D5F5E3"` (light green), `"D6EAF8"` (light blue)

### Fonts

Font sizes are set when creating paragraph styles:
- Title: `Pt(44)`
- Day headers: `Pt(14)`
- Day numbers: `Pt(18)`

Change these values to make text larger or smaller.

### Slide Size

The presentation dimensions are set in `main()`:
```python
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)
```

For letter-size landscape printing, you might use `Inches(11)` x `Inches(8.5)`.

## Troubleshooting

**Import errors**: Make sure `python-pptx` is installed: `pip3 install python-pptx`

**Wrong weekday alignment**: Verify `calendar.setfirstweekday(calendar.MONDAY)` is called before any `monthcalendar` calls.

**Colors not appearing**: The XML namespace in `set_cell_fill` must match exactly. Don't modify the namespace URL.

**Validation failures**: Check that `DAY_NAMES` order matches the `setfirstweekday` setting.
