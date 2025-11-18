# EOD2 - Shipment Dashboard

Automated daily shipment reporting system for processing and analyzing shipment data.

## Overview

This project generates comprehensive Excel dashboards from shipment CSV data files, providing daily insights into:
- VIN counts by customer and tag type
- Daily shipment metrics with percentage increases
- CarMax unique VIN tracking (New status, no tags)
- Top vehicles and tag distribution
- Distance analytics

## Features

### Main Dashboard (`shipment_dashboard_excel.py`)
- **Automatic CSV Detection**: Processes any CSV file in the directory
- **Dual File Processing**: 
  - Main EOD file for general shipments
  - EOD Update-2 file for CarMax unique VIN tracking
- **Multiple Worksheets**:
  1. Dashboard Summary - Key metrics overview
  2. Pivot Tables - Customer × Tag breakdown (all dates & today's data)
  3. Tag Distribution - Visual analysis with charts
  4. Top Vehicles - Most shipped vehicles with bar charts
  5. Raw Data - Complete filtered dataset

### Key Metrics Tracked
- Count of VIN by Customer and Tag Type (with date range)
- Count of VIN Created Today (with percentage increase)
- CarMax VINs with New Status (No Tags) - grouped by date
- Most shipped vehicles
- Average distance per shipment

### Data Filtering
- Automatically excludes orders tagged with "Quote"
- Filters CarMax orders by:
  - Customer = CarMax
  - Vehicle Status = New
  - Tags = Empty/Null

## Requirements

```bash
pip install pandas openpyxl
```

## Usage

1. Place your CSV files in the directory:
   - `MB EOD Update_*.csv` (main shipment data)
   - `MB EOD Update-2_*.csv` (CarMax tracking data)

2. Run the script:
```bash
python shipment_dashboard_excel.py
```

3. The script will:
   - Automatically detect the most recent CSV files
   - Process and filter the data
   - Generate `shipment_dashboard_YYYY-MM-DD.xlsx`

## Output

The generated Excel file includes:
- Professional formatting with color-coded sections
- Interactive charts (pie and bar)
- Alternating row colors for readability
- Properly sized columns for all data
- Total rows with highlighted backgrounds

## File Structure

```
eod2/
├── shipment_dashboard_excel.py    # Main Excel dashboard generator
├── shipment_dashboard_pdf.py      # PDF dashboard generator (legacy)
├── shipment_dashboard.py          # HTML dashboard generator (legacy)
├── .gitignore                     # Excludes CSV and Excel files
└── README.md                      # This file
```

## Notes

- CSV and Excel files are excluded from version control for data privacy
- The script processes data from the most recently modified CSV file
- All date ranges are automatically calculated from the data
- Percentage increases are calculated as (Today's Count / Total Count) × 100

## Author

Created for automated daily shipment reporting and analytics.

## License

Internal use only.

