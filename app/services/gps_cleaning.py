import pandas as pd
from pathlib import Path

NAME_SUFFIXES = {'jr', 'sr', 'ii', 'iii', 'iv'}


def normalize_paychex_folder(folder_path: str) -> pd.DataFrame:
    all_dfs = []
    folder = Path(folder_path)

    for file in folder.glob('*.xlsx'):
        df = pd.read_excel(file)

        df[['last_name', 'first_name']] = (
            df['Employee Name']
            .astype(str)
            .str.split(',', expand=True)
        )

        df['first_name'] = df['first_name'].str.strip().str.lower()
        df['last_name'] = df['last_name'].str.strip().str.lower()

        # Robust date
        df['date'] = pd.to_datetime(df['Date'], errors='coerce')
        df['date'] = df['date'].fillna(pd.to_datetime(df['Work Start'], errors='coerce'))
        df['date'] = df['date'].dt.normalize()

        df['start_paychex'] = pd.to_datetime(df['Work Start'], errors='coerce')
        df['end_paychex'] = pd.to_datetime(df['Work End'], errors='coerce')

        all_dfs.append(
            df[['first_name', 'last_name', 'date', 'start_paychex', 'end_paychex']]
        )

    df = pd.concat(all_dfs, ignore_index=True)

    # Merge duplicates
    df = df.groupby(['first_name', 'last_name', 'date'], as_index=False).agg({
        'start_paychex': 'min',
        'end_paychex': 'max'
    })

    return df


def normalize_samsara_folder(folder_path: str) -> pd.DataFrame:
    all_dfs = []
    folder = Path(folder_path)

    def extract_first_last(parts):
        if not parts:
            return pd.Series([None, None])
        if parts[-1] in NAME_SUFFIXES:
            parts = parts[:-1]
        if len(parts) == 1:
            return pd.Series([parts[0], None])
        return pd.Series([parts[0], parts[-1]])

    for file in folder.glob('*.xlsx'):
        xls = pd.ExcelFile(file)
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)

            name_parts = (
                df['Driver Name']
                .astype(str)
                .str.lower()
                .str.replace(r'[^\w\s]', '', regex=True)
                .str.split()
            )
            df[['first_name', 'last_name']] = name_parts.apply(extract_first_last)

            df['first_name'] = df['first_name'].str.strip()
            df['last_name'] = df['last_name'].str.strip()

            # Robust date
            df['date'] = pd.to_datetime(df['Start Date'], errors='coerce')
            df['date'] = df['date'].fillna(pd.to_datetime(df['End Date'], errors='coerce'))
            df['date'] = df['date'].dt.normalize()

            df['start_samsara'] = pd.to_datetime(
                df['Start Date'].astype(str) + ' ' + df['Start Time'].astype(str),
                errors='coerce'
            )
            df['end_samsara'] = pd.to_datetime(
                df['End Date'].astype(str) + ' ' + df['End Time'].astype(str),
                errors='coerce'
            )

            all_dfs.append(
                df[['first_name', 'last_name', 'date', 'start_samsara', 'end_samsara']]
            )

    df = pd.concat(all_dfs, ignore_index=True)

    # Merge duplicates
    df = df.groupby(['first_name', 'last_name', 'date'], as_index=False).agg({
        'start_samsara': 'min',
        'end_samsara': 'max'
    })

    return df
