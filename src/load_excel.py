"""Reusable, read-only helpers for inspecting raw Excel workbooks."""

from pathlib import Path
import re

import pandas as pd
from pandas.api.types import is_numeric_dtype


EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
KNOWN_PIO_FIELDS = {
    "PIS_MST_IVC_DT": "invoice date / sales posting date",
    "PIS_CMP_KND": "company or brand code",
    "PIS_SERI": "likely vehicle model/series code; validate variants before matching",
    "PIS_MDL_YY": "vehicle model year",
    "PIS_PNO": "accessory part number",
    "SumOfPIS_INST_QT": "total installed quantity",
    "SumOfPIS_CRP_CFM_PRI": "total confirmed installed price / sales amount (revenue)",
    "YYYYMM": "reporting month",
    "Model": "vehicle model name",
    "Part Description": "accessory / part description",
    "Deliminated Date": "derived/normalized date; exact meaning needs confirmation",
}
FIELD_KEYWORDS = {
    "date": ("date", "ivc_dt"),
    "month": ("month", "yyyymm"),
    "brand / company": ("brand", "company", "make", "cmp_knd"),
    "vehicle model": ("model",),
    "vehicle model/series code": ("series", "seri"),
    "vehicle model year": ("model year", "mdl_yy"),
    "part number / description": ("part", "accessory", "description", "pno"),
    "installed quantity / demand": ("qty", "quantity", "installed", "inst_qt"),
    "vehicle volume": ("volume", "wholesale", "fleet"),
    "revenue / sales dollars": ("revenue", "sales", "amount", "dollar", "price", "cfm_pri"),
}
MONTH_NAMES = ("jan", "feb", "mar", "apr", "may", "jun",
               "jul", "aug", "sep", "oct", "nov", "dec")
MONTH_LABELS = {
    "jan": "Jan", "feb": "Feb", "mar": "Mar", "apr": "Apr",
    "may": "May", "jun": "Jun", "jul": "Jul", "aug": "Aug",
    "sep": "Sep", "oct": "Oct", "nov": "Nov", "dec": "Dec",
}
BRAND_BLOCK_LABELS = {"hma", "gma", "kus"}
SUBTOTAL_LABELS = {
    "hma total", "gma total", "h+g total", "kus total", "grand total"
}


def find_excel_file(raw_dir="data/raw"):
    """Return the only Excel workbook in *raw_dir* or raise a clear error."""
    raw_path = Path(raw_dir)
    if not raw_path.exists():
        raise FileNotFoundError(f"Raw data directory does not exist: {raw_path.resolve()}")

    files = sorted(
        path for path in raw_path.iterdir()
        if (
            path.is_file()
            and path.suffix.lower() in EXCEL_EXTENSIONS
            and not path.name.startswith("~$")
        )
    )
    if not files:
        raise FileNotFoundError(
            f"No Excel workbook found in {raw_path.resolve()}. "
            "Expected a .xlsx, .xlsm, or .xls file."
        )
    if len(files) > 1:
        choices = "\n".join(f"  - {path.name}" for path in files)
        raise RuntimeError(
            "Multiple Excel workbooks were found. Choose one explicitly before "
            f"continuing:\n{choices}"
        )
    return files[0]


def list_sheet_names(file_path):
    """List workbook sheet names without changing the source workbook."""
    with pd.ExcelFile(file_path) as workbook:
        return workbook.sheet_names


def read_sheet_preview(file_path, sheet_name, nrows=20, header=None):
    """Read a small sheet preview. Raw inspection should use header=None."""
    return pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        nrows=nrows,
        header=header,
    )


def summarize_missing_values(df):
    """Return missing counts and percentages for non-empty columns."""
    if df.empty:
        return pd.DataFrame(columns=["column", "missing_count", "missing_percent"])

    summary = pd.DataFrame({
        "column": [str(column) for column in df.columns],
        "missing_count": df.isna().sum().values,
        "missing_percent": (df.isna().mean().values * 100).round(1),
    })
    return summary.sort_values(
        ["missing_count", "column"], ascending=[False, True]
    ).reset_index(drop=True)


def _searchable_labels(df):
    labels = [str(column).strip() for column in df.columns]
    # In raw mode, likely headers may be present within the first several rows.
    for value in df.head(10).to_numpy().ravel():
        if pd.notna(value):
            labels.append(str(value).strip())
    return labels


def identify_likely_useful_fields(df):
    """Identify candidate business fields from columns and early sheet rows."""
    labels = _searchable_labels(df)
    normalized_to_original = {_normalize_label(label): label for label in labels}
    matches = {}

    # Prefer the known PIO source fields and give each one a clear business meaning.
    for source_field, meaning in KNOWN_PIO_FIELDS.items():
        normalized = _normalize_label(source_field)
        if normalized in normalized_to_original:
            matches[meaning] = [normalized_to_original[normalized]]

    # Retain generic keyword matching for sheets whose field names differ.
    for concept, keywords in FIELD_KEYWORDS.items():
        found = sorted({
            label for label in labels
            if any(keyword in _normalize_label(label) for keyword in keywords)
        })
        if found:
            matches.setdefault(concept, [])
            matches[concept] = sorted(set(matches[concept] + found))
    return matches


def _find_column(df, expected_name):
    """Return the actual column matching *expected_name*, ignoring case/spacing."""
    expected = _normalize_label(expected_name)
    return next(
        (column for column in df.columns if _normalize_label(column) == expected),
        None,
    )


def _parse_date_column(series, column_name):
    """Parse a candidate date column without modifying the source values."""
    if _normalize_label(column_name) == "pis_mst_ivc_dt":
        text = series.astype("string").str.replace(r"\.0$", "", regex=True)
        return pd.to_datetime(text, format="%Y%m%d", errors="coerce")
    return pd.to_datetime(series, errors="coerce")


def create_pio_sales_profile(df):
    """Create lightweight, read-only EDA summaries for a header-based PIO table."""
    profile = {
        "candidate_columns": [str(column) for column in df.columns],
        "data_types": {
            str(column): str(dtype) for column, dtype in df.dtypes.items()
        },
        "missing_summary": summarize_missing_values(df),
        "duplicate_row_count": int(df.duplicated().sum()),
        "date_ranges": {},
        "yyyymm_range": None,
        "total_installed_quantity": None,
        "total_revenue": None,
        "unique_model_count": None,
        "unique_part_number_count": None,
        "warnings": [],
    }

    date_columns = [
        column for name in ("PIS_MST_IVC_DT", "Deliminated Date")
        if (column := _find_column(df, name)) is not None
    ]
    for column in date_columns:
        parsed = _parse_date_column(df[column], column)
        unparsed = int((df[column].notna() & parsed.isna()).sum())
        missing = int(df[column].isna().sum())
        profile["date_ranges"][str(column)] = {
            "min": parsed.min() if parsed.notna().any() else None,
            "max": parsed.max() if parsed.notna().any() else None,
            "unparsed_non_null_count": unparsed,
        }
        if missing:
            profile["warnings"].append(
                f"{column} has {missing:,} missing date value(s)."
            )
        if unparsed:
            profile["warnings"].append(
                f"{column} has {unparsed:,} non-empty value(s) that did not parse as dates."
            )

    yyyymm_column = _find_column(df, "YYYYMM")
    if yyyymm_column is not None:
        yyyymm = (
            df[yyyymm_column].astype("string").str.replace(r"\.0$", "", regex=True)
        )
        valid = yyyymm.where(yyyymm.str.fullmatch(r"\d{6}", na=False))
        profile["yyyymm_range"] = {
            "min": valid.min() if valid.notna().any() else None,
            "max": valid.max() if valid.notna().any() else None,
            "invalid_non_null_count": int((df[yyyymm_column].notna() & valid.isna()).sum()),
        }

    quantity_column = _find_column(df, "SumOfPIS_INST_QT")
    if quantity_column is not None:
        quantity = pd.to_numeric(df[quantity_column], errors="coerce")
        invalid = int((df[quantity_column].notna() & quantity.isna()).sum())
        profile["total_installed_quantity"] = quantity.sum(min_count=1)
        if not is_numeric_dtype(df[quantity_column]) or invalid:
            profile["warnings"].append(
                f"{quantity_column} is not fully numeric ({invalid:,} non-empty value(s) failed conversion)."
            )
        if (quantity < 0).any():
            profile["warnings"].append(
                f"{quantity_column} contains {int((quantity < 0).sum()):,} negative value(s)."
            )

    revenue_column = _find_column(df, "SumOfPIS_CRP_CFM_PRI")
    if revenue_column is not None:
        revenue = pd.to_numeric(df[revenue_column], errors="coerce")
        invalid = int((df[revenue_column].notna() & revenue.isna()).sum())
        profile["total_revenue"] = revenue.sum(min_count=1)
        if not is_numeric_dtype(df[revenue_column]) or invalid:
            profile["warnings"].append(
                f"{revenue_column} is not fully numeric ({invalid:,} non-empty value(s) failed conversion)."
            )
        if (revenue < 0).any():
            profile["warnings"].append(
                f"{revenue_column} contains {int((revenue < 0).sum()):,} negative value(s)."
            )

    model_column = _find_column(df, "Model")
    if model_column is not None:
        profile["unique_model_count"] = int(df[model_column].nunique(dropna=True))

    part_column = _find_column(df, "PIS_PNO")
    if part_column is not None:
        profile["unique_part_number_count"] = int(
            df[part_column].nunique(dropna=True)
        )

    if not profile["warnings"]:
        profile["warnings"].append(
            "No top-level date, numeric-type, or negative-value warnings detected."
        )
    return profile


def _normalize_label(value):
    """Normalize a cell label for conservative structural matching."""
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip()).lower()


def _matching_rows(df, predicate):
    matches = []
    for index, row in df.iterrows():
        labels = [_normalize_label(value) for value in row if pd.notna(value)]
        found = sorted({label for label in labels if predicate(label)})
        if found:
            matches.append({"row_number": int(index) + 1, "matched_values": found})
    return matches


def _excel_column_name(column_index):
    """Convert a zero-based column index to an Excel-style column label."""
    number = column_index + 1
    label = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        label = chr(65 + remainder) + label
    return label


def _year_from_value(value):
    label = _normalize_label(value)
    match = re.fullmatch(r"(2025|2026)(?:\.0)?", label)
    return int(match.group(1)) if match else None


def _month_from_value(value):
    label = _normalize_label(value)
    return MONTH_LABELS.get(label[:3]) if label[:3] in MONTH_LABELS else None


def _rows_within(matches, start_row, end_row):
    return [
        match for match in matches
        if start_row <= match["row_number"] <= end_row
    ]


def _year_month_columns_for_section(
    df, section_name, start_row, end_row, year_rows, month_rows
):
    """Pair repeated month columns with forward-filled year headers."""
    section_year_rows = _rows_within(year_rows, start_row, end_row)
    section_month_rows = _rows_within(month_rows, start_row, end_row)
    if not section_year_rows or not section_month_rows:
        return []

    candidates = []
    for month_match in section_month_rows:
        month_row_number = month_match["row_number"]
        year_match = min(
            section_year_rows,
            key=lambda item: abs(item["row_number"] - month_row_number),
        )
        year_row_number = year_match["row_number"]
        year_values = df.iloc[year_row_number - 1]
        month_values = df.iloc[month_row_number - 1]

        current_year = None
        for column_index in range(df.shape[1]):
            detected_year = _year_from_value(year_values.iloc[column_index])
            if detected_year is not None:
                current_year = detected_year
            month = _month_from_value(month_values.iloc[column_index])
            if month is not None:
                candidates.append({
                    "section": section_name,
                    "column_index": column_index,
                    "column_number": column_index + 1,
                    "excel_column": _excel_column_name(column_index),
                    "year": current_year,
                    "month": month,
                    "year_header_row": year_row_number,
                    "month_header_row": month_row_number,
                })
    return candidates


def detect_vehicle_report_structure(df):
    """Build a read-only layout map for the Wholesale/Fleet report."""
    wholesale_rows = _matching_rows(df, lambda label: "wholesale" in label)
    fleet_rows = _matching_rows(df, lambda label: "fleet" in label)
    brand_rows = _matching_rows(df, lambda label: label in BRAND_BLOCK_LABELS)
    subtotal_rows = _matching_rows(df, lambda label: label in SUBTOTAL_LABELS)
    year_rows = _matching_rows(
        df,
        lambda label: bool(re.fullmatch(r"2025(?:\.0)?|2026(?:\.0)?", label)),
    )
    month_rows = _matching_rows(df, lambda label: label[:3] in MONTH_NAMES)

    section_starts = []
    if wholesale_rows:
        section_starts.append(("Wholesale", wholesale_rows[0]["row_number"]))
    if fleet_rows:
        section_starts.append(("Fleet", fleet_rows[0]["row_number"]))
    section_starts.sort(key=lambda item: item[1])

    nonempty_rows = df.notna().any(axis=1)
    last_used_row = int(nonempty_rows[nonempty_rows].index.max()) + 1 if nonempty_rows.any() else 0
    sections = []
    all_year_month_columns = []
    for position, (section_name, start_row) in enumerate(section_starts):
        next_start = (
            section_starts[position + 1][1]
            if position + 1 < len(section_starts)
            else last_used_row + 1
        )
        end_row = max(start_row, next_start - 1)
        section_brands = _rows_within(brand_rows, start_row, end_row)
        section_subtotals = _rows_within(subtotal_rows, start_row, end_row)
        year_month_columns = _year_month_columns_for_section(
            df, section_name, start_row, end_row, year_rows, month_rows
        )
        all_year_month_columns.extend(year_month_columns)
        sections.append({
            "section": section_name,
            "start_row": start_row,
            "end_row": end_row,
            "row_range": f"{start_row}-{end_row}",
            "brand_blocks": section_brands,
            "subtotal_rows": section_subtotals,
            "year_header_rows": _rows_within(year_rows, start_row, end_row),
            "month_header_rows": _rows_within(month_rows, start_row, end_row),
            "year_month_columns": year_month_columns,
            "notes": [
                f"Parse the {section_name} section separately.",
                "Treat detected section boundaries as estimates until visually confirmed.",
                "Exclude subtotal rows from model-level records unless a later use explicitly requires totals.",
            ],
        })

    return {
        "possible_wholesale_section_rows": wholesale_rows,
        "possible_fleet_section_rows": fleet_rows,
        "possible_brand_block_rows": brand_rows,
        "possible_subtotal_rows": subtotal_rows,
        "possible_year_header_rows": year_rows,
        "possible_month_header_rows": month_rows,
        "has_wide_year_month_columns": len(all_year_month_columns) >= 4,
        "year_month_column_candidates": all_year_month_columns,
        "sections": sections,
        "cleaning_readiness_notes": [
            "Wholesale and Fleet should be parsed separately.",
            "Subtotal rows should likely be excluded from model-level cleaning.",
            "Year/month headers must be reconstructed before wide-to-long reshaping.",
        ],
    }


def infer_sheet_format(df):
    """Heuristically classify a sheet as long-format, wide-format, or report-style."""
    if df.empty:
        return "unknown/empty"

    early_rows = df.head(10)
    non_null_by_row = early_rows.notna().sum(axis=1)
    month_pattern = re.compile(
        r"^(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)", re.IGNORECASE
    )
    values = [str(value).strip() for value in early_rows.to_numpy().ravel() if pd.notna(value)]
    repeated_month_labels = sum(bool(month_pattern.match(value)) for value in values)
    title_or_sparse_rows = int((non_null_by_row <= max(2, df.shape[1] * 0.25)).sum())

    if title_or_sparse_rows >= 2 or repeated_month_labels >= 4:
        return "report-style"
    if df.shape[1] >= 12 and repeated_month_labels >= 2:
        return "wide-format"
    return "long-format (likely)"


def _cleaning_notes(raw_df, sheet_name, inferred_format):
    notes = []
    lower_name = sheet_name.lower()
    if raw_df.empty:
        return ["Sheet is empty."]
    if raw_df.isna().mean().mean() > 0.30:
        notes.append("High overall missingness; merged cells or layout spacing may be present.")
    if inferred_format == "report-style":
        notes.append("Inspect title rows, merged cells, multi-row headers, and section boundaries.")
    if "vehicle_wholesale" in lower_name:
        notes.extend([
            "Treat this sheet as a report containing Wholesale and Fleet sections, not as a clean table.",
            "Parse Wholesale and Fleet separately using visually confirmed section boundaries.",
            "Do not treat row 1 as the header until Wholesale and Fleet sections are located.",
            "Header=0 labels such as 'Unnamed: 0' are layout artifacts, not usable cleaned column names.",
            "Reconstruct year/month headers before wide-to-long reshaping.",
            "Likely exclude subtotal and grand-total rows from model-level cleaning to avoid double counting.",
        ])
    if "pio_sales" in lower_name:
        notes.append(
            "Candidate long table: validate dates, identifiers, quantities, and sales values "
            "before later aggregation."
        )
    if not notes:
        notes.append("Confirm the true header row, data types, and any totals before cleaning.")
    return notes


def inspect_sheet(file_path, sheet_name):
    """Inspect a sheet and always return the same public dictionary keys."""
    raw_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    header_df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
    lower_name = sheet_name.lower()
    is_pio_sales = "pio_sales" in lower_name
    is_vehicle_report = "vehicle_wholesale" in sheet_name.lower()
    inferred_format = "report-style" if is_vehicle_report else infer_sheet_format(raw_df)
    report_structure = (
        detect_vehicle_report_structure(raw_df) if is_vehicle_report else None
    )
    pio_profile = create_pio_sales_profile(header_df) if is_pio_sales else None

    return {
        "sheet_name": sheet_name,
        "shape": raw_df.shape,
        "raw_preview": raw_df.head(20),
        "raw_column_names": list(raw_df.columns),
        "header_column_names": [str(column) for column in header_df.columns],
        "missing_summary": summarize_missing_values(raw_df),
        "inferred_format": inferred_format,
        "useful_fields": identify_likely_useful_fields(header_df),
        "cleaning_notes": _cleaning_notes(raw_df, sheet_name, inferred_format),
        "report_structure": report_structure,
        "pio_profile": pio_profile,
    }
