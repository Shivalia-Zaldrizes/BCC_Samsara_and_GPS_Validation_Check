# app/services/time_aggregation.py

import pandas as pd

def aggregate_all_events(df: pd.DataFrame, _=None) -> pd.DataFrame:
    for col in ['paychex_start','paychex_end','samsara_start','samsara_end']:
        if col not in df.columns:
            df[col] = pd.NaT
        else:
            df[col] = pd.to_datetime(df[col], errors='coerce')

    # Compute hours
    df['paychex_hours'] = (df['paychex_end'] - df['paychex_start']).dt.total_seconds() / 3600
    df['samsara_hours'] = (df['samsara_end'] - df['samsara_start']).dt.total_seconds() / 3600
    return df
