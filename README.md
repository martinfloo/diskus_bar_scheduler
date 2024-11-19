# Bar Shift Scheduler

An automated scheduling system for managing bar staff shifts, processing availability data, and generating organized Excel schedules.

## Features

- Automated shift assignment based on staff availability
- Color-coded Excel schedule output
- Smart handling of shift requirements per day
- Automatic name matching with manual review options
- Weekend and holiday handling
- Staff workload balancing

## Requirements

- Python 3.6+
- Required packages:
  ```
  pandas
  openpyxl
  ```

## Setup

1. Create a directory for your project and ensure you have the following files:
   - `bar_scheduler.py` (main script)
   - `members.txt` (list of staff members, one per line)
   - Availability survey responses in CSV format

2. Install required packages:
   ```bash
   pip install pandas openpyxl
   ```

3. Configure the script by updating these variables in `BarScheduler` class:
   ```python
   YEAR = 2024                # Current year
   MONTH = 11                 # Month number (1-12)
   USERPATH = "/your/path/"   # Directory path for input/output files
   ```

## Shift Types

| Shift    | Time         | Color in Excel |
|----------|-------------|----------------|
| Opening  | 12:30-17:00 | Pink          |
| Middle   | 16:50-20:30 | Blue          |
| Closing  | 20:20-00:30 | Green         |

Note: No closing shifts are scheduled on Mondays.

## Schedule Rules

- Each staff member gets approximately 2 shifts per week
- No consecutive day shifts for the same person
- Staff requirements vary by day:
  - Monday: 2 opening, 2 middle, no closing
  - Tuesday-Friday: Varies between 2-3 staff per shift
- Weekends are marked in gray and not scheduled
- Non-responding staff are marked in dark gray

## File Requirements

### members.txt
```
John Smith
Jane Doe
[one name per line]
```

### Availability CSV Format
The CSV should contain:
- A column named "Navn og etternavn" for staff names
- Date columns with format "[Day]. [Month abbrev]"
- Shift availability in each cell

## Usage

1. Ensure all required files are in place
2. Run the script:
   ```bash
   python bar_scheduler.py
   ```

The script will generate an Excel file named `{month}_schedule_{year}.xlsx` in your specified directory.

## Output Format

The generated Excel file includes:
- Color-coded shift schedule
- Shift totals per person
- Availability tracking
- Weekend highlighting
- Manual review sheet for name matching issues
- Daily shift count totals

## Contributing

Feel free to submit issues and enhancement requests!
