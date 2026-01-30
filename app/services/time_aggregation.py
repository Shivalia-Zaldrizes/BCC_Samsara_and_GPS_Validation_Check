import pandas as pd

def aggregate_all_events(paychex_df: pd.DataFrame, samsara_df: pd.DataFrame) -> pd.DataFrame:
    df = pd.merge(
        paychex_df,
        samsara_df,
        on=['first_name', 'last_name', 'date'],
        how='outer'
    )

    # Compute differences and categories
    df['start_diff'] = (df['start_paychex'] - df['start_samsara']).dt.total_seconds() / 60
    df['end_diff'] = (df['end_paychex'] - df['end_samsara']).dt.total_seconds() / 60

    def categorize(diff_minutes):
        if pd.isna(diff_minutes):
          return 'Missing'
        abs_diff = abs(diff_minutes)
        if abs_diff < 15:
            return 'Within Reason'
        elif 15 <= abs_diff < 30:
            return 'Within Reason'
        elif 30 <= abs_diff < 60:
            return 'Slight Difference'
        elif 60 <= abs_diff < 120:
            return 'Large Difference'
        else:  # 120 minutes or more
            return 'Outrageous'

    df['start_category'] = df['start_diff'].apply(categorize)
    df['end_category'] = df['end_diff'].apply(categorize)

    return df
