# Score Report Tool
A command-line tool to generate Excel score reports from CSV files.
It calculates statistics, highlights key metrics, and optionally generates charts.

## Features
- Read scores from CSV
- Generate formatted Excel report
- Summary statistics (Average, Median, Pass Rate, etc.)
- Conditional formatting (Pass Rate / Median alerts)
- Optional Pass/Fail chart
- Robust error handling

## Requirements
- Python 3.9+
- openpyxl

## Usage
Basic usage:
```bash
python report.py scores.csv report.xlsx
```
Custom pass score:
```bash
python report.py scores.csv report.xlsx --pass-score 75
```
Generate chart:
```bash
python report.py scores.csv report.xlsx --chart
```

## Input CSV format
```csv
name,score
Alice,85
Bob,72
Charlie,90
Eunice,91
```

## Output
Excel report with Summary and Details sheets

Automatic formatting and conditional coloring
