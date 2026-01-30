from app.services.gps_cleaning import normalize_paychex_folder, normalize_samsara_folder
from app.services.time_aggregation import aggregate_all_events
from app.services.excel_export import export_weekly_report

PAYCHEX_EXCEL_PATH = r'C:\Users\AliciaH\OneDrive - buckleycable.com\Documents\Samsara and GPS Validation Check Project\Samsara & GPS Validation\Import Files\Excel\Paychex_Files'
SAMSARA_EXCEL_PATH = r'C:\Users\AliciaH\OneDrive - buckleycable.com\Documents\Samsara and GPS Validation Check Project\Samsara & GPS Validation\Import Files\Excel\Samsara_Files'
OUTPUT_EXCEL_PATH = r'C:\Users\AliciaH\OneDrive - buckleycable.com\Documents\Samsara and GPS Validation Check Project\Samsara & GPS Validation\Export Files\Excel\Weekly_Report.xlsx'

def main():
    print("Reading Paychex files...")
    paychex_df = normalize_paychex_folder(PAYCHEX_EXCEL_PATH)
    if paychex_df.empty:
        print("No Paychex data found. Exiting.")
        return

    print("Reading Samsara files...")
    samsara_df = normalize_samsara_folder(SAMSARA_EXCEL_PATH)
    if samsara_df.empty:
        print("No Samsara data found. Exiting.")
        return

    print("Aggregating events...")
    agg_df = aggregate_all_events(paychex_df, samsara_df)

    print("Exporting weekly report...")
    export_weekly_report(agg_df, OUTPUT_EXCEL_PATH)
    print(f"Export complete: {OUTPUT_EXCEL_PATH}")


if __name__ == "__main__":
    main()
