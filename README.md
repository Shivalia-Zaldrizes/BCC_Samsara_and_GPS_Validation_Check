# BCC_Samsara_and_GPS_Validation_Check
Checks the start/stop time of GPS data and compares to clock in and out times submitted by Paychex to then highlight cells with discrepancies.


qb_excel_sync/
├── app/
│   ├── services/
│   │   ├── data_cleaning.py        # Cleans QB / Paychex payroll CSVs
│   │   ├── excel_hours.py          # Normalizes Excel sheets (names as columns)
│   │   ├── gps_cleaning.py         # Normalizes Samsara / GPS clock data
│   │   ├── time_aggregation.py     # Handles multiple punches per day
│   │   ├── discrepancy.py          # Time difference + status classification
│   │   ├── excel_local.py          # Writes output Excel locally
│   │
│   ├── config.py                   # File paths, feature flags
│   ├── logging.py                  # App logger
│   ├── main.py                     # Orchestrates the full pipeline
│
├── Export Files/
│   ├── CSV/                        # Raw vendor exports (QB / Paychex / Samsara)
│   └── Excel/                      # Optional intermediate exports
│
├── Import Files/
│   ├── GPS/                        # Samsara / GPS raw sheets
│   ├── Payroll/                    # QB / Paychex CSVs
│
├── Output Files/
│   └── payroll_audit.xlsx          # Final discrepancy report
│
├── logs/
│   └── qb_excel_sync.log
│
├── venv/
│
├── .env.example.txt
├── requirements.txt
