import pandas as pd
from pathlib import Path
import re

def _safe_col(df: pd.DataFrame, candidates: list[str]):
    """
    Returns the first column name found in df from the candidates list.
    """
    for col in candidates:
        if col in df.columns:
            return col
    return None


def _parse_date_safe(series):
    return pd.to_datetime(series, errors="coerce")

def _normalize_name(series: pd.Series) -> pd.Series:
    return (
        series.astype(str)
        .str.upper()
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )

def read_paychex_files(folder: Path) -> pd.DataFrame:
    rows = []

    for file in folder.glob("*.xlsx"):
        try:
            df = pd.read_excel(file, engine="openpyxl")
        except PermissionError:
            print(f"[SKIPPED] File is open or locked: {file.name}")
            continue

        df.columns = df.columns.str.strip()

        name_col = _safe_col(df, ["Name"])
        date_col = _safe_col(df, ["Date"])
        start_col = _safe_col(df, ["Work Start"])
        end_col = _safe_col(df, ["Work End"])

        if not name_col or not date_col:
            print(f"[WARN] Missing required columns in {file.name}")
            continue

        names = df[name_col].astype(str).str.split(",", expand=True)
        last = names[0].str.strip()
        first = names[1].str.strip() if names.shape[1] > 1 else ""

        temp = pd.DataFrame({
            "first_name": _normalize_name(first),
            "last_name": _normalize_name(last),
            "date": _parse_date_safe(df[date_col]).dt.date,
            "paychex_start": _parse_date_safe(df[start_col]) if start_col else pd.NaT,
            "paychex_end": _parse_date_safe(df[end_col]) if end_col else pd.NaT,
        })

        rows.append(temp)

    if not rows:
        return pd.DataFrame(columns=[
            "first_name", "last_name", "date",
            "paychex_start", "paychex_end"
        ])

    return pd.concat(rows, ignore_index=True)


def read_samsara_files(folder: Path) -> pd.DataFrame:
    dfs = []

    for file in folder.glob("*.xlsx"):
        df = pd.read_excel(file, engine="openpyxl")
        df.columns = df.columns.str.strip()

        name_col = df.get("Driver Name")
        date_col = df.get("Start Date")
        start_col = df.get("Start Time")
        end_col = df.get("End Time")

        if name_col is None or date_col is None:
            print(f"[WARN] Missing required columns in {file.name}")
            continue

        temp = pd.DataFrame({
            "date": pd.to_datetime(date_col, errors="coerce").dt.date,
            "start_time": start_col,
            "end_time": end_col
        })

        split = name_col.astype(str).str.strip().str.split(" ", n=1, expand=True)
        temp["first_name"] = _normalize_name(split[0])
        temp["last_name"] = _normalize_name(split[1]) if split.shape[1] > 1 else ""

        # Combine date + time into full datetimes
        temp["start_samsara"] = pd.to_datetime(
            temp["date"].astype(str) + " " + temp["start_time"].astype(str),
            errors="coerce"
        )

        temp["end_samsara"] = pd.to_datetime(
            temp["date"].astype(str) + " " + temp["end_time"].astype(str),
            errors="coerce"
        )

        # Drop raw time columns
        temp = temp.drop(columns=["start_time", "end_time"])

        dfs.append(temp)

        print(f"[INFO] Loaded Samsara file: {file.name} ({len(temp)} rows)")

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()


def merge_paychex_samsara(paychex, samsara):
    return pd.merge(
        paychex,
        samsara,
        on=["first_name", "last_name", "date"],
        how="outer"
    )