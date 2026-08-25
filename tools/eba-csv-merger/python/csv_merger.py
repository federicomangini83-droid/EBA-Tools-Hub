"""EBA CSV Merger processing engine for Pyodide."""

import io
import json
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd

MAPPING_SHEET_NAME = "Mapping_dpm2_df_All"


def _read_csv_robust(file_path: Path) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    encodings = ["utf-8-sig", "utf-8", "latin-1", "cp1252"]
    separators = [",", ";", "\t"]
    last_error = None

    for encoding in encodings:
        for separator in separators:
            try:
                dataframe = pd.read_csv(
                    file_path,
                    encoding=encoding,
                    sep=separator,
                    low_memory=False,
                )
                if dataframe.shape[1] == 1:
                    continue
                dataframe.insert(0, "source_file", file_path.name)
                return dataframe, f"encoding={encoding}, separator={repr(separator)}"
            except Exception as exc:
                last_error = str(exc)

    return None, last_error


def _json_safe_records(dataframe: pd.DataFrame, limit: int = 200) -> List[dict]:
    preview = dataframe.head(limit).copy()
    preview = preview.where(pd.notna(preview), None)
    return preview.to_dict(orient="records")


def process_csv_files(file_paths_json: str, mapping_path: str) -> str:
    file_paths = [Path(path) for path in json.loads(file_paths_json)]
    logs = ["=== EBA CSV Merger started ==="]
    valid_dataframes = []
    skipped_files = []

    for file_path in file_paths:
        dataframe, read_info = _read_csv_robust(file_path)
        if dataframe is None:
            skipped_files.append(file_path.name)
            logs.append(f"Skipped: {file_path.name} (unsupported encoding or separator)")
            continue

        valid_dataframes.append(dataframe)
        logs.append(
            f"Loaded: {file_path.name} | {read_info} | "
            f"rows={len(dataframe):,} | columns={len(dataframe.columns):,}"
        )

    if not valid_dataframes:
        raise ValueError("No valid CSV files could be read.")

    logs.append(f"Merging {len(valid_dataframes)} valid files...")
    merged = pd.concat(valid_dataframes, axis=0, join="outer", ignore_index=True)
    rows_before_filter = len(merged)

    logs.append(f"Loading mapping sheet: {MAPPING_SHEET_NAME}...")
    mapping = pd.read_excel(
        mapping_path,
        sheet_name=MAPPING_SHEET_NAME,
        engine="openpyxl",
        usecols=["FR_TABLE_L", "FR_ROW", "FR_COLUMN", "datapoint"],
    )

    required_input = {"source_file", "datapoint"}
    missing_input = sorted(required_input - set(merged.columns))
    if missing_input:
        raise ValueError(
            "The merged CSV data is missing required columns: "
            + ", ".join(missing_input)
        )

    merged["source_file"] = (
        merged["source_file"]
        .astype("string")
        .str.replace(r"\.csv$", "", regex=True, case=False)
    )
    merged = merged.dropna(subset=["datapoint"])

    merged = merged.merge(
        mapping,
        how="left",
        left_on=["source_file", "datapoint"],
        right_on=["FR_TABLE_L", "datapoint"],
    )

    drop_candidates = ["templateID", "reported", "name", "value"]
    merged = merged.drop(
        columns=[column for column in drop_candidates if column in merged.columns]
    )

    merged = merged.rename(
        columns={
            "FR_TABLE_L": "IdTab",
            "FR_ROW": "Row",
            "FR_COLUMN": "Column",
            "datapoint": "Datapoint",
            "source_file": "IdTab_file",
        }
    )

    first_columns = [
        "IdTab_file",
        "IdTab",
        "Row",
        "Column",
        "Datapoint",
        "factValue",
    ]
    first_columns = [column for column in first_columns if column in merged.columns]
    merged = merged[
        first_columns
        + [column for column in merged.columns if column not in first_columns]
    ]

    sort_columns = [
        column for column in ["IdTab", "Row", "Column"] if column in merged.columns
    ]
    if sort_columns:
        merged = merged.sort_values(by=sort_columns, na_position="last")

    mapped_rows = int(merged["IdTab"].notna().sum()) if "IdTab" in merged.columns else 0
    unmapped_rows = len(merged) - mapped_rows

    csv_buffer = io.StringIO(newline="")
    merged.to_csv(csv_buffer, index=False, lineterminator="\n")

    logs.extend(
        [
            f"Rows read before filtering: {rows_before_filter:,}",
            f"Rows in final output: {len(merged):,}",
            f"Rows matched to mapping: {mapped_rows:,}",
            f"Rows without mapping coordinates: {unmapped_rows:,}",
            f"Files processed: {len(valid_dataframes)}",
            f"Files skipped: {len(skipped_files)}",
            "=== Processing completed ===",
        ]
    )

    result = {
        "csv": csv_buffer.getvalue(),
        "preview": _json_safe_records(merged),
        "columns": list(merged.columns),
        "row_count": int(len(merged)),
        "mapped_rows": mapped_rows,
        "unmapped_rows": unmapped_rows,
        "processed_files": len(valid_dataframes),
        "skipped_files": skipped_files,
        "logs": logs,
    }
    return json.dumps(result, ensure_ascii=False, default=str)
