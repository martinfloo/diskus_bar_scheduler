# Bar Shift Scheduler

Ever tried managing a bar with almost 90 employers, who all got their own preferences for when to work?

bar_scheduler is my solution for the tedious manual work. 

On average creating the shift list took 5-6 hours - now it only takes 1 second. 

Get all your employers to answer for when to work in a google forms document. I have given a mock example in this repo. 

Feel free to use my implemenation if you need it!

## Requirements
- Python 3.6+
- `pandas`, `openpyxl`

## Setup
1. Required files:
   - `bar_scheduler.py`
   - `members.txt` (staff list)
   - Availability CSV

2. Installation:
```bash
pip install pandas openpyxl
```

## Configuration
Set in `BarScheduler` class:
```python
YEAR = 2024
MONTH = 11
USERPATH = "/your/path/"
```

## Shifts
| Shift    | Time         | Staff Requirements |
|----------|-------------|-------------------|
| Opening  | 12:30-17:00 | 2 per day      |
| Middle   | 16:50-20:30 | 2-3 per day      |
| Closing* | 20:20-00:30 | 2-3 per day      |
*No closing shifts on Mondays

## Key Features
- Automated shift assignment
- Color-coded Excel output
- Staff workload balancing
- Weekend/holiday handling
- Name matching with review options

## Usage
```bash
python bar_scheduler.py
```
Generates `{month}_schedule_{year}.xlsx`

## File Formats

### members.txt
```
Name1
Name2
```

### CSV Requirements
- Column: "Navn og etternavn" (names)
- Date columns: "[Day]. [Month abbrev]"
- Shift availability in cells

## Output
- Color-coded schedule
- Shift/availability tracking
- Weekend highlighting
- Name matching review sheet
- Daily totals