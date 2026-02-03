from pathlib import Path
from app.services.gps_cleaning import read_paychex_files, read_samsara_files, merge_paychex_samsara
from app.services.time_aggregation import aggregate_all_events
from app.services.excel_export import export_weekly_report

# Input / Output paths
PAYCHEX_EXCEL_PATH = Path(
    r"C:\Users\AliciaH\OneDrive - buckleycable.com\Documents\Samsara and GPS Validation Check Project\Samsara & GPS Validation\Import Files\Excel\Paychex_Files"
)
SAMSARA_EXCEL_PATH = Path(
    r"C:\Users\AliciaH\OneDrive - buckleycable.com\Documents\Samsara and GPS Validation Check Project\Samsara & GPS Validation\Import Files\Excel\Samsara_Files"
)
OUTPUT_EXCEL_PATH = Path(
    r"C:\Users\AliciaH\OneDrive - buckleycable.com\Documents\Samsara and GPS Validation Check Project\Samsara & GPS Validation\Export Files\Excel"
)


def main():
    print("Reading Paychex files...")
    paychex_df = read_paychex_files(PAYCHEX_EXCEL_PATH)
    if paychex_df.empty:
        print("No Paychex data found. Exiting.")
        return
    print(f"[INFO] Paychex rows loaded: {len(paychex_df)}")

    print("Reading Samsara files...")
    samsara_df = read_samsara_files(SAMSARA_EXCEL_PATH)
    if samsara_df.empty:
        print("No Samsara data found. Exiting.")
        return
    print(f"[INFO] Samsara rows loaded: {len(samsara_df)}")

    print("Merging datasets...")
    merged_df = merge_paychex_samsara(paychex_df, samsara_df)
    print(f"[INFO] Total merged rows: {len(merged_df)}")

    print("Aggregating events...")
    agg_df = aggregate_all_events(merged_df)
    print(f"[INFO] Aggregated rows: {len(agg_df)}")

    print("Exporting weekly report...")
    export_weekly_report(agg_df, OUTPUT_EXCEL_PATH)
    print(f"[INFO] Export complete: {OUTPUT_EXCEL_PATH}")


if __name__ == "__main__":
    main()
