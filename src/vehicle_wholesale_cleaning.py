"""Helpers for parsing the report-style ``Vehicle_Wholesale_Data`` sheet."""

from pathlib import Path
import re

import pandas as pd

from src.load_excel import detect_vehicle_report_structure, find_excel_file


SHEET_NAME = "Vehicle_Wholesale_Data"
BRAND_GROUPS = {"hma": "HMA", "gma": "GMA", "kus": "KUS"}
TOTAL_LABELS = {
    "hma total": ("HMA", "HMA Total"),
    "gma total": ("GMA", "GMA Total"),
    "h+g total": ("H+G", "H+G Total"),
    "kus total": ("KUS", "KUS Total"),
    "grand total": ("Grand Total", "Grand Total"),
}
MONTH_NUMBERS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4,
    "may": 5, "jun": 6, "jul": 7, "aug": 8,
    "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


def load_vehicle_wholesale(raw_dir="data/raw"):
    """Find the workbook and read the vehicle report in raw mode."""
    workbook_path = find_excel_file(raw_dir)
    raw_df = pd.read_excel(
        workbook_path,
        sheet_name=SHEET_NAME,
        header=None,
    )
    return workbook_path, raw_df


def _normalize(value):
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip()).lower()


def _as_text(value):
    """Preserve identifiers without adding Excel-style trailing decimals."""
    if pd.isna(value):
        return pd.NA
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    return text if text else pd.NA


def _column_with_label(raw_df, row_number, expected_label):
    row = raw_df.iloc[row_number - 1]
    expected = _normalize(expected_label)
    matches = [
        column_index
        for column_index, value in row.items()
        if _normalize(value) == expected
    ]
    if len(matches) != 1:
        raise ValueError(
            f"Expected one '{expected_label}' column on Excel row "
            f"{row_number}, found {len(matches)}."
        )
    return matches[0]


def _brand_column_for_section(raw_df, section):
    candidates = []
    for match in section["brand_blocks"]:
        row_index = match["row_number"] - 1
        for column_index, value in raw_df.iloc[row_index].items():
            if _normalize(value) in BRAND_GROUPS:
                candidates.append(column_index)
    if not candidates or len(set(candidates)) != 1:
        raise ValueError(
            f"Could not determine one brand-group column for "
            f"{section['section']}."
        )
    return candidates[0]


def build_year_month_column_map(raw_df, report_structure=None):
    """Reconstruct monthly columns, using cross-section alignment when needed."""
    structure = report_structure or detect_vehicle_report_structure(raw_df)
    sections = structure.get("sections", [])
    if not sections:
        raise ValueError("No Wholesale/Fleet sections were detected.")

    candidates = [
        candidate
        for section in sections
        for candidate in section["year_month_columns"]
    ]
    if not candidates:
        raise ValueError("No year/month column candidates were detected.")

    # A blank/merged year label in one section can be inferred from the same
    # physical Excel column in another section, provided there is no conflict.
    year_by_column = {}
    month_by_column = {}
    for candidate in candidates:
        column_index = candidate["column_index"]
        year = candidate["year"]
        month = candidate["month"]
        if year is not None:
            existing_year = year_by_column.get(column_index)
            if existing_year is not None and existing_year != year:
                raise ValueError(
                    f"Conflicting years detected for column {column_index}: "
                    f"{existing_year} and {year}."
                )
            year_by_column[column_index] = int(year)
        existing_month = month_by_column.get(column_index)
        if existing_month is not None and existing_month != month:
            raise ValueError(
                f"Conflicting months detected for column {column_index}: "
                f"{existing_month} and {month}."
            )
        month_by_column[column_index] = month

    records = []
    for candidate in candidates:
        column_index = candidate["column_index"]
        year = candidate["year"]
        inferred_year = year is None
        resolved_year = year_by_column.get(column_index)
        if resolved_year is None:
            raise ValueError(
                f"Year could not be resolved for Excel column "
                f"{candidate['excel_column']}."
            )
        month_label = month_by_column[column_index]
        month_number = MONTH_NUMBERS[_normalize(month_label)[:3]]
        records.append({
            "channel": candidate["section"],
            "column_index": column_index,
            "source_column": candidate["excel_column"],
            "year": resolved_year,
            "month_label": month_label,
            "month": pd.Timestamp(resolved_year, month_number, 1),
            "year_header_row": candidate["year_header_row"],
            "month_header_row": candidate["month_header_row"],
            "year_was_inferred": inferred_year,
        })

    column_map = (
        pd.DataFrame(records)
        .drop_duplicates(["channel", "column_index"])
        .sort_values(["channel", "month"])
        .reset_index(drop=True)
    )
    if column_map.duplicated(["channel", "month"]).any():
        raise ValueError("Duplicate channel-month columns were reconstructed.")
    return column_map


def clean_vehicle_wholesale(raw_df, report_structure=None):
    """Convert Wholesale and Fleet report sections to an auditable long table."""
    structure = report_structure or detect_vehicle_report_structure(raw_df)
    column_map = build_year_month_column_map(raw_df, structure)
    records = []

    for section in structure["sections"]:
        channel = section["section"]
        section_columns = column_map[column_map["channel"] == channel]
        if section_columns.empty:
            raise ValueError(f"No monthly columns resolved for {channel}.")

        month_header_rows = section["month_header_rows"]
        if not month_header_rows:
            raise ValueError(f"No month header row detected for {channel}.")
        header_row = month_header_rows[0]["row_number"]
        brand_column = _brand_column_for_section(raw_df, section)
        model_column = _column_with_label(raw_df, header_row, "Model")
        model_code_column = _column_with_label(
            raw_df, header_row, "Model Code"
        )

        current_brand = None
        first_data_row = header_row + 1
        for source_row in range(first_data_row, section["end_row"] + 1):
            row = raw_df.iloc[source_row - 1]
            row_label = _normalize(row.iloc[brand_column])
            if row_label in BRAND_GROUPS:
                current_brand = BRAND_GROUPS[row_label]

            is_total_row = row_label in TOTAL_LABELS
            model = _as_text(row.iloc[model_column])
            model_code = _as_text(row.iloc[model_code_column])
            if is_total_row:
                brand_group, total_label = TOTAL_LABELS[row_label]
                model = pd.NA
                model_code = pd.NA
            elif pd.notna(model):
                brand_group = current_brand
                total_label = pd.NA
                if brand_group is None:
                    raise ValueError(
                        f"Model row {source_row} has no detected brand group."
                    )
            else:
                # Ignore spacing, title, and other non-data rows.
                continue

            for column in section_columns.itertuples(index=False):
                source_value = row.iloc[column.column_index]
                numeric_value = pd.to_numeric(
                    pd.Series([source_value]), errors="coerce"
                ).iloc[0]
                if pd.isna(source_value):
                    volume_status = "missing"
                elif pd.isna(numeric_value):
                    volume_status = "nonnumeric"
                else:
                    volume_status = "numeric"

                records.append({
                    "month": column.month,
                    "channel": channel,
                    "brand_group": brand_group,
                    "model": model,
                    "model_code": model_code,
                    "vehicle_volume": numeric_value,
                    "source_row": source_row,
                    "source_column": column.source_column,
                    "is_total_row": is_total_row,
                    "total_label": total_label,
                    "volume_status": volume_status,
                })

    columns = [
        "month", "channel", "brand_group", "model", "model_code",
        "vehicle_volume", "source_row", "source_column", "is_total_row",
        "total_label", "volume_status",
    ]
    cleaned = pd.DataFrame(records, columns=columns)
    cleaned["month"] = pd.to_datetime(cleaned["month"])
    cleaned["channel"] = cleaned["channel"].astype("string")
    cleaned["brand_group"] = cleaned["brand_group"].astype("string")
    cleaned["model"] = cleaned["model"].astype("string")
    cleaned["model_code"] = cleaned["model_code"].astype("string")
    cleaned["vehicle_volume"] = cleaned["vehicle_volume"].astype("Float64")
    cleaned["source_row"] = cleaned["source_row"].astype("Int64")
    cleaned["source_column"] = cleaned["source_column"].astype("string")
    cleaned["is_total_row"] = cleaned["is_total_row"].astype(bool)
    cleaned["total_label"] = cleaned["total_label"].astype("string")
    cleaned["volume_status"] = cleaned["volume_status"].astype("string")
    return cleaned.sort_values(
        ["channel", "source_row", "month"]
    ).reset_index(drop=True)


def split_model_and_total_rows(cleaned_df):
    """Return separate model-level and subtotal/total long tables."""
    model_rows = cleaned_df.loc[~cleaned_df["is_total_row"]].copy()
    total_rows = cleaned_df.loc[cleaned_df["is_total_row"]].copy()
    return model_rows, total_rows


def save_vehicle_processed_tables(
    model_rows,
    total_rows,
    processed_dir="data/processed",
):
    """Save model-level and total-row tables as separate ignored CSV files."""
    output_dir = Path(processed_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = {
        "vehicle_model_month": output_dir / "vehicle_model_month.csv",
        "vehicle_total_month": output_dir / "vehicle_total_month.csv",
    }
    model_rows.to_csv(
        paths["vehicle_model_month"],
        index=False,
        date_format="%Y-%m-%d",
    )
    total_rows.to_csv(
        paths["vehicle_total_month"],
        index=False,
        date_format="%Y-%m-%d",
    )
    return paths


def vehicle_cleaning_summaries(cleaned_df):
    """Create compact validation summaries without exporting data."""
    model_rows, _ = split_model_and_total_rows(cleaned_df)

    row_counts = (
        cleaned_df.groupby(
            ["channel", "brand_group", "is_total_row"],
            dropna=False,
            as_index=False,
        )
        .agg(
            long_row_count=("source_row", "size"),
            source_row_count=("source_row", "nunique"),
        )
        .sort_values(["channel", "is_total_row", "brand_group"])
    )

    monthly_volume = (
        model_rows.groupby(["month", "channel"], as_index=False)
        .agg(
            vehicle_volume=(
                "vehicle_volume",
                lambda values: values.sum(min_count=1),
            ),
            model_count=("model", "nunique"),
            missing_volume_count=(
                "volume_status",
                lambda values: int((values == "missing").sum()),
            ),
            nonnumeric_volume_count=(
                "volume_status",
                lambda values: int((values == "nonnumeric").sum()),
            ),
        )
        .sort_values(["month", "channel"])
    )

    total_vs_model = (
        cleaned_df.groupby(["channel", "is_total_row"], as_index=False)
        .agg(
            long_row_count=("source_row", "size"),
            source_row_count=("source_row", "nunique"),
        )
        .sort_values(["channel", "is_total_row"])
    )

    volume_quality = (
        cleaned_df.groupby(
            ["channel", "is_total_row", "volume_status"],
            as_index=False,
        )
        .size()
        .rename(columns={"size": "row_count"})
        .sort_values(["channel", "is_total_row", "volume_status"])
    )

    months_covered = (
        cleaned_df.groupby("channel", as_index=False)
        .agg(
            first_month=("month", "min"),
            last_month=("month", "max"),
            month_count=("month", "nunique"),
        )
        .sort_values("channel")
    )

    return {
        "row_counts_by_channel_brand": row_counts.reset_index(drop=True),
        "monthly_vehicle_volume_by_channel": monthly_volume.reset_index(
            drop=True
        ),
        "total_rows_vs_model_rows": total_vs_model.reset_index(drop=True),
        "volume_quality": volume_quality.reset_index(drop=True),
        "months_covered": months_covered.reset_index(drop=True),
    }
