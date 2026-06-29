"""Exploratory PIO-to-vehicle matching and KPI helpers.

This module uses the cleaned in-memory outputs from steps 02 and 03. It does
not export data and does not build forecasting models.
"""

import re
import unicodedata

import pandas as pd

from src.load_excel import detect_vehicle_report_structure
from src.pio_sales_eda import (
    MONTH,
    QUANTITY,
    REVENUE,
    load_pio_sales,
    prepare_pio_sales_for_eda,
)
from src.vehicle_wholesale_cleaning import (
    clean_vehicle_wholesale,
    load_vehicle_wholesale,
    split_model_and_total_rows,
)


GENESIS_MODEL_PATTERN = re.compile(
    r"^(g70|g80|g90|gv60|gv70|gv80)(?:\s|$)"
)
DEFAULT_MODEL_ALIASES = {
    # PIO uses GV60 while the vehicle report labels the model GV60 EV.
    "gv60": "gv60 ev",
}
MATCH_KEYS = ["month", "tentative_brand_group", "normalized_model"]


def load_matching_inputs(raw_dir="data/raw"):
    """Load PIO data and the cleaned vehicle model rows using prior helpers."""
    workbook_path, pio_raw = load_pio_sales(raw_dir)
    pio_analysis = prepare_pio_sales_for_eda(pio_raw)

    vehicle_workbook_path, vehicle_raw = load_vehicle_wholesale(raw_dir)
    if vehicle_workbook_path.resolve() != workbook_path.resolve():
        raise ValueError("PIO and vehicle data resolved to different workbooks.")
    vehicle_structure = detect_vehicle_report_structure(vehicle_raw)
    vehicle_long = clean_vehicle_wholesale(vehicle_raw, vehicle_structure)
    vehicle_models, vehicle_totals = split_model_and_total_rows(vehicle_long)

    return {
        "workbook_path": workbook_path,
        "pio_raw": pio_raw,
        "pio_analysis": pio_analysis,
        "vehicle_long": vehicle_long,
        "vehicle_model_rows": vehicle_models,
        "vehicle_total_rows": vehicle_totals,
    }


def normalize_model_name(value, aliases=None):
    """Normalize a model label conservatively for exact key matching."""
    if pd.isna(value):
        return pd.NA
    text = unicodedata.normalize("NFKD", str(value))
    text = text.encode("ascii", "ignore").decode("ascii").lower().strip()
    text = text.replace("&", " and ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return pd.NA
    alias_map = DEFAULT_MODEL_ALIASES if aliases is None else aliases
    return alias_map.get(text, text)


def tentative_pio_brand_group(pio_brand_code, normalized_model):
    """Apply the explicitly tentative H/K to HMA/GMA/KUS mapping."""
    code = "" if pd.isna(pio_brand_code) else str(pio_brand_code).strip().upper()
    model = "" if pd.isna(normalized_model) else str(normalized_model)
    if code == "K":
        return "KUS"
    if code == "H":
        return "GMA" if GENESIS_MODEL_PATTERN.match(model) else "HMA"
    return "Unmapped"


def aggregate_pio_model_month(pio_analysis):
    """Aggregate PIO quantity and revenue to one tentative model-month key."""
    data = pio_analysis.copy()
    data["normalized_model"] = data["Model"].map(normalize_model_name)
    data["tentative_brand_group"] = [
        tentative_pio_brand_group(code, model)
        for code, model in zip(
            data["PIS_CMP_KND"], data["normalized_model"]
        )
    ]

    result = (
        data.groupby(
            [
                MONTH,
                "PIS_CMP_KND",
                "tentative_brand_group",
                "normalized_model",
            ],
            dropna=False,
            as_index=False,
        )
        .agg(
            model=("Model", _join_unique_text),
            pio_quantity=(QUANTITY, "sum"),
            pio_revenue=(REVENUE, "sum"),
        )
        .rename(columns={
            MONTH: "month",
            "PIS_CMP_KND": "pio_brand_code",
        })
    )
    return result[
        [
            "month", "pio_brand_code", "model", "pio_quantity",
            "pio_revenue", "normalized_model", "tentative_brand_group",
        ]
    ].sort_values(MATCH_KEYS).reset_index(drop=True)


def aggregate_pio_model_part_month(pio_analysis):
    """Aggregate PIO quantity/revenue to model-part-month grain."""
    data = pio_analysis.copy()
    data["normalized_model"] = data["Model"].map(normalize_model_name)
    data["tentative_brand_group"] = [
        tentative_pio_brand_group(code, model)
        for code, model in zip(
            data["PIS_CMP_KND"], data["normalized_model"]
        )
    ]

    result = (
        data.groupby(
            [
                MONTH,
                "PIS_CMP_KND",
                "tentative_brand_group",
                "normalized_model",
                "PIS_PNO",
                "Part Description",
            ],
            dropna=False,
            as_index=False,
        )
        .agg(
            model=("Model", _join_unique_text),
            quantity=(QUANTITY, "sum"),
            revenue=(REVENUE, "sum"),
        )
        .rename(columns={
            MONTH: "month",
            "PIS_CMP_KND": "pio_brand_code",
            "PIS_PNO": "part_number",
            "Part Description": "part_description",
        })
    )
    return result[
        [
            "month", "pio_brand_code", "model", "part_number",
            "part_description", "quantity", "revenue",
            "normalized_model", "tentative_brand_group",
        ]
    ].sort_values(
        MATCH_KEYS + ["part_number"]
    ).reset_index(drop=True)


def _join_unique_text(values):
    unique = sorted({
        str(value).strip()
        for value in values
        if pd.notna(value) and str(value).strip()
    })
    return " | ".join(unique) if unique else pd.NA


def _normalize_model_code(value):
    """Normalize Excel model-code values without changing their meaning."""
    if pd.isna(value):
        return pd.NA
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip().upper()
    return text if text else pd.NA


def _join_unique_codes(values):
    unique = sorted({
        str(value)
        for value in values
        if pd.notna(value) and str(value).strip()
    })
    return " | ".join(unique) if unique else pd.NA


def _sum_with_missing(values):
    return values.sum(min_count=1)


def aggregate_vehicle_model_month(vehicle_model_rows):
    """Aggregate cleaned model rows and pivot Wholesale/Fleet volumes."""
    data = vehicle_model_rows.copy()
    data["normalized_model"] = data["model"].map(normalize_model_name)
    data["tentative_brand_group"] = data["brand_group"].astype("string")

    grouped = (
        data.groupby(
            MATCH_KEYS + ["channel"],
            dropna=False,
            as_index=False,
        )
        .agg(
            vehicle_model=("model", _join_unique_text),
            vehicle_model_code=("model_code", _join_unique_text),
            vehicle_volume=("vehicle_volume", _sum_with_missing),
        )
    )

    metadata = (
        grouped.groupby(MATCH_KEYS, dropna=False, as_index=False)
        .agg(
            vehicle_model=("vehicle_model", _join_unique_text),
            vehicle_model_code=("vehicle_model_code", _join_unique_text),
        )
    )
    volumes = (
        grouped.pivot(
            index=MATCH_KEYS,
            columns="channel",
            values="vehicle_volume",
        )
        .reset_index()
        .rename(columns={
            "Wholesale": "wholesale_volume",
            "Fleet": "fleet_volume",
        })
    )
    for column in ("wholesale_volume", "fleet_volume"):
        if column not in volumes:
            volumes[column] = pd.NA

    result = metadata.merge(volumes, on=MATCH_KEYS, how="outer")
    result["wholesale_volume"] = pd.to_numeric(
        result["wholesale_volume"], errors="coerce"
    ).astype("Float64")
    result["fleet_volume"] = pd.to_numeric(
        result["fleet_volume"], errors="coerce"
    ).astype("Float64")
    # Conservative exploratory total: require both channel values.
    result["total_vehicle_volume"] = (
        result["wholesale_volume"] + result["fleet_volume"]
    ).astype("Float64")
    result["total_vehicle_volume_is_exploratory"] = True
    return result.sort_values(MATCH_KEYS).reset_index(drop=True)


def model_code_variant_diagnostics(pio_raw, vehicle_model_rows):
    """Compare PIO PIS_SERI with vehicle model-code/variant relationships."""
    pio = pio_raw[
        ["PIS_CMP_KND", "Model", "PIS_SERI"]
    ].drop_duplicates().copy()
    pio["normalized_model"] = pio["Model"].map(normalize_model_name)
    pio["pio_series_code"] = pio["PIS_SERI"].map(
        _normalize_model_code
    )
    pio["tentative_brand_group"] = [
        tentative_pio_brand_group(code, model)
        for code, model in zip(
            pio["PIS_CMP_KND"], pio["normalized_model"]
        )
    ]

    vehicle = vehicle_model_rows[
        ["brand_group", "model", "model_code"]
    ].drop_duplicates().copy()
    vehicle["normalized_vehicle_model"] = vehicle["model"].map(
        normalize_model_name
    )
    vehicle["vehicle_model_code"] = vehicle["model_code"].map(
        _normalize_model_code
    )
    vehicle = vehicle.rename(columns={
        "brand_group": "tentative_brand_group",
        "model": "vehicle_model",
    })

    pio_model_codes = (
        pio.groupby(
            [
                "PIS_CMP_KND", "tentative_brand_group",
                "Model", "normalized_model",
            ],
            dropna=False,
            as_index=False,
        )
        .agg(
            pio_series_count=("pio_series_code", "nunique"),
            pio_series_values=("pio_series_code", _join_unique_codes),
        )
        .rename(columns={
            "PIS_CMP_KND": "pio_brand_code",
            "Model": "pio_model",
        })
    )
    pio_multiple_series = pio_model_codes.loc[
        pio_model_codes["pio_series_count"] > 1
    ].sort_values(
        ["tentative_brand_group", "pio_model"]
    ).reset_index(drop=True)

    vehicle_code_models = (
        vehicle.groupby(
            ["tentative_brand_group", "vehicle_model_code"],
            dropna=False,
            as_index=False,
        )
        .agg(
            vehicle_model_count=("vehicle_model", "nunique"),
            vehicle_models=("vehicle_model", _join_unique_text),
        )
    )
    vehicle_multiple_models = vehicle_code_models.loc[
        vehicle_code_models["vehicle_model_count"] > 1
    ].sort_values(
        ["tentative_brand_group", "vehicle_model_code"]
    ).reset_index(drop=True)

    vehicle_by_exact_name = (
        vehicle.groupby(
            ["tentative_brand_group", "normalized_vehicle_model"],
            dropna=False,
            as_index=False,
        )
        .agg(
            vehicle_models=("vehicle_model", _join_unique_text),
            vehicle_model_codes=(
                "vehicle_model_code", _join_unique_codes
            ),
        )
        .rename(columns={
            "normalized_vehicle_model": "normalized_model"
        })
    )
    exact_name_comparison = pio_model_codes.merge(
        vehicle_by_exact_name,
        on=["tentative_brand_group", "normalized_model"],
        how="left",
        validate="one_to_one",
    )

    def has_code_overlap(row):
        if (
            pd.isna(row["pio_series_values"])
            or pd.isna(row["vehicle_model_codes"])
        ):
            return False
        pio_codes = set(str(row["pio_series_values"]).split(" | "))
        vehicle_codes = set(
            str(row["vehicle_model_codes"]).split(" | ")
        )
        return bool(pio_codes & vehicle_codes)

    exact_name_comparison["code_overlap"] = exact_name_comparison.apply(
        has_code_overlap, axis=1
    )
    exact_name_comparison = exact_name_comparison.sort_values(
        ["tentative_brand_group", "pio_model"]
    ).reset_index(drop=True)

    variant_candidates = pio.merge(
        vehicle,
        left_on=["tentative_brand_group", "pio_series_code"],
        right_on=["tentative_brand_group", "vehicle_model_code"],
        how="inner",
        validate="many_to_many",
    )
    variant_candidates = variant_candidates.loc[
        variant_candidates["normalized_model"].ne(
            variant_candidates["normalized_vehicle_model"]
        )
    ].copy()
    variant_candidates["vehicle_name_has_variant_label"] = (
        variant_candidates["vehicle_model"]
        .astype("string")
        .str.contains(
            r"\b(?:HEV|PHEV|EV|Coupe|N)\b",
            case=False,
            regex=True,
        )
        .fillna(False)
    )
    variant_candidates = (
        variant_candidates[[
            "PIS_CMP_KND", "tentative_brand_group", "Model",
            "pio_series_code", "vehicle_model", "vehicle_model_code",
            "vehicle_name_has_variant_label",
        ]]
        .rename(columns={
            "PIS_CMP_KND": "pio_brand_code",
            "Model": "pio_model",
        })
        .drop_duplicates()
        .sort_values(
            [
                "tentative_brand_group", "pio_model",
                "pio_series_code", "vehicle_model",
            ]
        )
        .reset_index(drop=True)
    )

    return {
        "pio_model_code_map": pio_model_codes.reset_index(drop=True),
        "pio_models_multiple_series": pio_multiple_series,
        "vehicle_code_model_map": vehicle_code_models.reset_index(
            drop=True
        ),
        "vehicle_codes_multiple_models": vehicle_multiple_models,
        "exact_name_code_comparison": exact_name_comparison,
        "variant_candidates": variant_candidates,
    }


def key_uniqueness_diagnostics(pio_model_month, vehicle_model_month):
    """Summarize duplicate match keys before attempting the merge."""
    records = []
    for name, data in (
        ("PIO model-month", pio_model_month),
        ("Vehicle model-month", vehicle_model_month),
    ):
        duplicate_mask = data.duplicated(MATCH_KEYS, keep=False)
        duplicate_keys = data.loc[duplicate_mask, MATCH_KEYS].drop_duplicates()
        records.append({
            "dataset": name,
            "row_count": int(len(data)),
            "duplicate_key_rows": int(duplicate_mask.sum()),
            "duplicate_key_count": int(len(duplicate_keys)),
            "many_to_many_warning": bool(duplicate_mask.any()),
        })
    return pd.DataFrame(records)


def add_exploratory_kpis(matched):
    """Calculate ratios only where the corresponding denominator is positive."""
    result = matched.copy()
    wholesale_valid = result["wholesale_volume"] > 0
    total_valid = result["total_vehicle_volume"] > 0
    result["penetration_wholesale"] = (
        result["pio_quantity"].div(result["wholesale_volume"])
        .where(wholesale_valid)
    )
    result["penetration_total"] = (
        result["pio_quantity"].div(result["total_vehicle_volume"])
        .where(total_valid)
    )
    result["PMVW"] = (
        result["pio_revenue"].div(result["wholesale_volume"])
        .where(wholesale_valid)
    )
    result["PMV_total"] = (
        result["pio_revenue"].div(result["total_vehicle_volume"])
        .where(total_valid)
    )
    return result


def match_pio_to_vehicle(pio_model_month, vehicle_model_month):
    """Outer-match unique model-month keys and add exploratory KPIs."""
    uniqueness = key_uniqueness_diagnostics(
        pio_model_month, vehicle_model_month
    )
    if uniqueness["many_to_many_warning"].any():
        raise ValueError(
            "Duplicate model-month keys detected; refusing a potentially "
            "many-to-many merge."
        )

    matched = pio_model_month.merge(
        vehicle_model_month,
        on=MATCH_KEYS,
        how="outer",
        validate="one_to_one",
        indicator=True,
    )
    matched["key_matched"] = matched["_merge"].eq("both")
    return add_exploratory_kpis(matched), uniqueness


def match_pio_parts_to_vehicle(pio_model_part_month, vehicle_model_month):
    """Attach unique vehicle model-month volumes to PIO model-part rows."""
    if vehicle_model_month.duplicated(MATCH_KEYS).any():
        raise ValueError("Vehicle model-month keys are not unique.")
    return pio_model_part_month.merge(
        vehicle_model_month,
        on=MATCH_KEYS,
        how="left",
        validate="many_to_one",
        indicator=True,
    )


def _safe_share(numerator, denominator):
    if pd.isna(denominator) or denominator == 0:
        return None
    return float(numerator / denominator)


def _coverage_table(pio_side, group_columns):
    data = pio_side.copy()
    data["_matched_quantity"] = data["pio_quantity"].where(
        data["key_matched"], 0
    )
    data["_matched_revenue"] = data["pio_revenue"].where(
        data["key_matched"], 0
    )
    data["_valid_wholesale_quantity"] = data["pio_quantity"].where(
        data["wholesale_volume"] > 0, 0
    )
    data["_valid_wholesale_revenue"] = data["pio_revenue"].where(
        data["wholesale_volume"] > 0, 0
    )
    data["_valid_total_quantity"] = data["pio_quantity"].where(
        data["total_vehicle_volume"] > 0, 0
    )
    data["_valid_total_revenue"] = data["pio_revenue"].where(
        data["total_vehicle_volume"] > 0, 0
    )

    result = (
        data.groupby(group_columns, dropna=False, as_index=False)
        .agg(
            pio_rows=("pio_quantity", "size"),
            matched_rows=("key_matched", "sum"),
            pio_quantity=("pio_quantity", "sum"),
            matched_quantity=("_matched_quantity", "sum"),
            pio_revenue=("pio_revenue", "sum"),
            matched_revenue=("_matched_revenue", "sum"),
            valid_wholesale_quantity=(
                "_valid_wholesale_quantity", "sum"
            ),
            valid_wholesale_revenue=(
                "_valid_wholesale_revenue", "sum"
            ),
            valid_total_quantity=("_valid_total_quantity", "sum"),
            valid_total_revenue=("_valid_total_revenue", "sum"),
        )
    )
    result["matched_row_share"] = (
        result["matched_rows"].div(result["pio_rows"])
    )
    result["matched_quantity_share"] = (
        result["matched_quantity"].div(result["pio_quantity"])
    )
    result["matched_revenue_share"] = (
        result["matched_revenue"].div(result["pio_revenue"])
    )
    result["valid_wholesale_quantity_share"] = (
        result["valid_wholesale_quantity"].div(result["pio_quantity"])
    )
    result["valid_wholesale_revenue_share"] = (
        result["valid_wholesale_revenue"].div(result["pio_revenue"])
    )
    result["valid_total_quantity_share"] = (
        result["valid_total_quantity"].div(result["pio_quantity"])
    )
    result["valid_total_revenue_share"] = (
        result["valid_total_revenue"].div(result["pio_revenue"])
    )
    return result


def matching_diagnostics(matched, uniqueness):
    """Build coverage and unmatched-key diagnostics."""
    pio_side = matched.loc[matched["_merge"].ne("right_only")].copy()
    vehicle_side = matched.loc[matched["_merge"].ne("left_only")].copy()
    matched_pio = pio_side["key_matched"]

    pio_quantity = pio_side["pio_quantity"].sum()
    pio_revenue = pio_side["pio_revenue"].sum()
    wholesale_valid = pio_side["wholesale_volume"] > 0
    total_valid = pio_side["total_vehicle_volume"] > 0
    coverage_summary = pd.DataFrame([{
        "pio_model_month_rows": int(len(pio_side)),
        "matched_model_month_rows": int(matched_pio.sum()),
        "matched_row_share": _safe_share(
            matched_pio.sum(), len(pio_side)
        ),
        "pio_quantity": pio_quantity,
        "matched_pio_quantity": pio_side.loc[
            matched_pio, "pio_quantity"
        ].sum(),
        "matched_quantity_share": _safe_share(
            pio_side.loc[matched_pio, "pio_quantity"].sum(),
            pio_quantity,
        ),
        "pio_revenue": pio_revenue,
        "matched_pio_revenue": pio_side.loc[
            matched_pio, "pio_revenue"
        ].sum(),
        "matched_revenue_share": _safe_share(
            pio_side.loc[matched_pio, "pio_revenue"].sum(),
            pio_revenue,
        ),
        "positive_wholesale_denominator_rows": int(
            wholesale_valid.sum()
        ),
        "positive_wholesale_denominator_row_share": _safe_share(
            wholesale_valid.sum(), len(pio_side)
        ),
        "wholesale_denominator_quantity_share": _safe_share(
            pio_side.loc[wholesale_valid, "pio_quantity"].sum(),
            pio_quantity,
        ),
        "wholesale_denominator_revenue_share": _safe_share(
            pio_side.loc[wholesale_valid, "pio_revenue"].sum(),
            pio_revenue,
        ),
        "positive_total_denominator_rows": int(
            total_valid.sum()
        ),
        "positive_total_denominator_row_share": _safe_share(
            total_valid.sum(), len(pio_side)
        ),
        "total_denominator_quantity_share": _safe_share(
            pio_side.loc[total_valid, "pio_quantity"].sum(),
            pio_quantity,
        ),
        "total_denominator_revenue_share": _safe_share(
            pio_side.loc[total_valid, "pio_revenue"].sum(),
            pio_revenue,
        ),
    }])
    coverage_table = pd.DataFrame([
        {
            "coverage_type": "Key match",
            "basis": "PIO model-month rows",
            "covered": int(matched_pio.sum()),
            "total": int(len(pio_side)),
            "coverage_share": _safe_share(
                matched_pio.sum(), len(pio_side)
            ),
        },
        {
            "coverage_type": "Key match",
            "basis": "PIO quantity",
            "covered": pio_side.loc[matched_pio, "pio_quantity"].sum(),
            "total": pio_quantity,
            "coverage_share": _safe_share(
                pio_side.loc[matched_pio, "pio_quantity"].sum(),
                pio_quantity,
            ),
        },
        {
            "coverage_type": "Key match",
            "basis": "PIO revenue",
            "covered": pio_side.loc[matched_pio, "pio_revenue"].sum(),
            "total": pio_revenue,
            "coverage_share": _safe_share(
                pio_side.loc[matched_pio, "pio_revenue"].sum(),
                pio_revenue,
            ),
        },
        {
            "coverage_type": "Positive Wholesale denominator",
            "basis": "PIO model-month rows",
            "covered": int(wholesale_valid.sum()),
            "total": int(len(pio_side)),
            "coverage_share": _safe_share(
                wholesale_valid.sum(), len(pio_side)
            ),
        },
        {
            "coverage_type": "Positive Wholesale denominator",
            "basis": "PIO quantity",
            "covered": pio_side.loc[
                wholesale_valid, "pio_quantity"
            ].sum(),
            "total": pio_quantity,
            "coverage_share": _safe_share(
                pio_side.loc[wholesale_valid, "pio_quantity"].sum(),
                pio_quantity,
            ),
        },
        {
            "coverage_type": "Positive Wholesale denominator",
            "basis": "PIO revenue",
            "covered": pio_side.loc[
                wholesale_valid, "pio_revenue"
            ].sum(),
            "total": pio_revenue,
            "coverage_share": _safe_share(
                pio_side.loc[wholesale_valid, "pio_revenue"].sum(),
                pio_revenue,
            ),
        },
        {
            "coverage_type": "Positive exploratory total denominator",
            "basis": "PIO model-month rows",
            "covered": int(total_valid.sum()),
            "total": int(len(pio_side)),
            "coverage_share": _safe_share(
                total_valid.sum(), len(pio_side)
            ),
        },
        {
            "coverage_type": "Positive exploratory total denominator",
            "basis": "PIO quantity",
            "covered": pio_side.loc[total_valid, "pio_quantity"].sum(),
            "total": pio_quantity,
            "coverage_share": _safe_share(
                pio_side.loc[total_valid, "pio_quantity"].sum(),
                pio_quantity,
            ),
        },
        {
            "coverage_type": "Positive exploratory total denominator",
            "basis": "PIO revenue",
            "covered": pio_side.loc[total_valid, "pio_revenue"].sum(),
            "total": pio_revenue,
            "coverage_share": _safe_share(
                pio_side.loc[total_valid, "pio_revenue"].sum(),
                pio_revenue,
            ),
        },
    ])

    unmatched_pio = pio_side.loc[
        ~pio_side["key_matched"],
        [
            "month", "pio_brand_code", "tentative_brand_group",
            "model", "normalized_model", "pio_quantity", "pio_revenue",
        ],
    ].sort_values(["month", "model"])

    unmatched_vehicle = vehicle_side.loc[
        vehicle_side["_merge"].eq("right_only"),
        [
            "month", "tentative_brand_group", "vehicle_model",
            "normalized_model", "wholesale_volume", "fleet_volume",
            "total_vehicle_volume",
        ],
    ].sort_values(["month", "vehicle_model"])
    pio_first_month = pio_side["month"].min()
    pio_last_month = pio_side["month"].max()
    within_pio_range = unmatched_vehicle["month"].between(
        pio_first_month, pio_last_month, inclusive="both"
    )
    unmatched_vehicle_within = unmatched_vehicle.loc[
        within_pio_range
    ].copy()
    unmatched_vehicle_outside = unmatched_vehicle.loc[
        ~within_pio_range
    ].copy()
    unmatched_vehicle_range_summary = pd.DataFrame([
        {
            "date_range_group": "Within PIO date range",
            "row_count": int(len(unmatched_vehicle_within)),
            "unique_models": int(
                unmatched_vehicle_within["normalized_model"].nunique()
            ),
            "first_month": unmatched_vehicle_within["month"].min(),
            "last_month": unmatched_vehicle_within["month"].max(),
        },
        {
            "date_range_group": "Outside PIO date range / future",
            "row_count": int(len(unmatched_vehicle_outside)),
            "unique_models": int(
                unmatched_vehicle_outside["normalized_model"].nunique()
            ),
            "first_month": unmatched_vehicle_outside["month"].min(),
            "last_month": unmatched_vehicle_outside["month"].max(),
        },
    ])
    unmatched_pio_by_model = (
        unmatched_pio.groupby(
            [
                "pio_brand_code", "tentative_brand_group",
                "model", "normalized_model",
            ],
            dropna=False,
            as_index=False,
        )
        .agg(
            unmatched_months=("month", "nunique"),
            pio_quantity=("pio_quantity", "sum"),
            pio_revenue=("pio_revenue", "sum"),
        )
        .sort_values("pio_revenue", ascending=False)
    )

    return {
        "coverage_summary": coverage_summary,
        "coverage_table": coverage_table,
        "coverage_by_month": _coverage_table(pio_side, ["month"]),
        "coverage_by_model": _coverage_table(
            pio_side,
            ["tentative_brand_group", "normalized_model", "model"],
        ),
        "unmatched_pio_model_months": unmatched_pio.reset_index(drop=True),
        "unmatched_vehicle_model_months": unmatched_vehicle.reset_index(
            drop=True
        ),
        "unmatched_pio_by_model": unmatched_pio_by_model.reset_index(
            drop=True
        ),
        "unmatched_vehicle_within_pio_range": (
            unmatched_vehicle_within.reset_index(drop=True)
        ),
        "unmatched_vehicle_outside_pio_range": (
            unmatched_vehicle_outside.reset_index(drop=True)
        ),
        "unmatched_vehicle_range_summary": (
            unmatched_vehicle_range_summary
        ),
        "key_uniqueness": uniqueness.reset_index(drop=True),
    }


def monthly_pio_vehicle_summary(pio_model_month, vehicle_model_month):
    """Compare monthly PIO totals with exploratory vehicle channel totals."""
    pio = (
        pio_model_month.groupby("month", as_index=False)
        .agg(
            pio_quantity=("pio_quantity", "sum"),
            pio_revenue=("pio_revenue", "sum"),
            pio_model_count=("normalized_model", "nunique"),
        )
    )
    vehicle = (
        vehicle_model_month.groupby("month", as_index=False)
        .agg(
            wholesale_volume=("wholesale_volume", _sum_with_missing),
            fleet_volume=("fleet_volume", _sum_with_missing),
            vehicle_model_count=("normalized_model", "nunique"),
            wholesale_missing_models=(
                "wholesale_volume",
                lambda values: int(values.isna().sum()),
            ),
            fleet_missing_models=(
                "fleet_volume",
                lambda values: int(values.isna().sum()),
            ),
        )
    )
    vehicle["total_vehicle_volume"] = (
        vehicle["wholesale_volume"] + vehicle["fleet_volume"]
    )
    vehicle["total_vehicle_volume_is_exploratory"] = True
    return pio.merge(vehicle, on="month", how="outer").sort_values("month")


def model_kpi_summary(matched):
    """Create denominator-aligned model-level exploratory KPI summaries."""
    data = matched.loc[matched["_merge"].eq("both")].copy()
    wholesale_valid = data["wholesale_volume"] > 0
    total_valid = data["total_vehicle_volume"] > 0
    data["_pio_quantity_wholesale"] = data["pio_quantity"].where(
        wholesale_valid
    )
    data["_pio_revenue_wholesale"] = data["pio_revenue"].where(
        wholesale_valid
    )
    data["_pio_quantity_total"] = data["pio_quantity"].where(total_valid)
    data["_pio_revenue_total"] = data["pio_revenue"].where(total_valid)
    data["_wholesale_valid"] = data["wholesale_volume"].where(
        wholesale_valid
    )
    data["_total_valid"] = data["total_vehicle_volume"].where(total_valid)

    result = (
        data.groupby(
            ["tentative_brand_group", "normalized_model", "model"],
            dropna=False,
            as_index=False,
        )
        .agg(
            matched_months=("month", "nunique"),
            pio_quantity=("pio_quantity", "sum"),
            pio_revenue=("pio_revenue", "sum"),
            wholesale_basis_quantity=(
                "_pio_quantity_wholesale", _sum_with_missing
            ),
            wholesale_basis_revenue=(
                "_pio_revenue_wholesale", _sum_with_missing
            ),
            wholesale_volume=("_wholesale_valid", _sum_with_missing),
            total_basis_quantity=(
                "_pio_quantity_total", _sum_with_missing
            ),
            total_basis_revenue=("_pio_revenue_total", _sum_with_missing),
            total_vehicle_volume=("_total_valid", _sum_with_missing),
        )
    )
    result["penetration_wholesale"] = (
        result["wholesale_basis_quantity"]
        .div(result["wholesale_volume"])
        .where(result["wholesale_volume"] > 0)
    )
    result["penetration_total"] = (
        result["total_basis_quantity"]
        .div(result["total_vehicle_volume"])
        .where(result["total_vehicle_volume"] > 0)
    )
    result["PMVW"] = (
        result["wholesale_basis_revenue"]
        .div(result["wholesale_volume"])
        .where(result["wholesale_volume"] > 0)
    )
    result["PMV_total"] = (
        result["total_basis_revenue"]
        .div(result["total_vehicle_volume"])
        .where(result["total_vehicle_volume"] > 0)
    )
    return result.sort_values(
        "pio_revenue", ascending=False
    ).reset_index(drop=True)


def kpi_outlier_candidates(model_summary, metric, lower=0.05, upper=0.95):
    """Flag low/high model candidates using simple empirical quantiles."""
    valid = model_summary.dropna(subset=[metric]).copy()
    if valid.empty:
        return valid.assign(outlier_direction=pd.Series(dtype="string"))
    low_threshold = valid[metric].quantile(lower)
    high_threshold = valid[metric].quantile(upper)
    low = valid.loc[valid[metric] <= low_threshold].copy()
    low["outlier_direction"] = "low"
    high = valid.loc[valid[metric] >= high_threshold].copy()
    high["outlier_direction"] = "high"
    return (
        pd.concat([low, high], ignore_index=True)
        .sort_values(metric)
        .reset_index(drop=True)
    )
