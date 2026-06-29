"""Read-only helpers for exploratory analysis of ``PIO_Sales_Data``."""

import pandas as pd

from src.load_excel import find_excel_file


PIO_SHEET_NAME = "PIO_Sales_Data"
PIO_COLUMNS = {
    "PIS_MST_IVC_DT": "invoice date / sales posting date",
    "PIS_CMP_KND": "company or brand code; observed values appear to be H and K",
    "PIS_SERI": "likely vehicle model/series code; useful for mapping validation but not sufficient alone",
    "PIS_MDL_YY": "vehicle model year",
    "PIS_PNO": "accessory part number",
    "SumOfPIS_INST_QT": "total installed quantity",
    "SumOfPIS_CRP_CFM_PRI": "total confirmed installed price / sales amount, treated as revenue",
    "YYYYMM": "reporting month",
    "Model": "vehicle model name",
    "Part Description": "accessory / part description",
    "Deliminated Date": "derived/normalized date field; exact meaning needs confirmation",
}

INVOICE_DATE = "_eda_invoice_date"
ALTERNATE_DATE = "_eda_deliminated_date"
MONTH = "_eda_month"
QUANTITY = "_eda_quantity"
REVENUE = "_eda_revenue"
UNIT_PRICE = "_eda_unit_price"


def load_pio_sales(raw_dir="data/raw"):
    """Find the raw workbook and read only ``PIO_Sales_Data``."""
    workbook_path = find_excel_file(raw_dir)
    data = pd.read_excel(workbook_path, sheet_name=PIO_SHEET_NAME, header=0)
    validate_pio_columns(data)
    return workbook_path, data


def validate_pio_columns(df):
    """Raise a clear error when an expected PIO field is absent."""
    missing = [column for column in PIO_COLUMNS if column not in df.columns]
    if missing:
        raise KeyError(
            "PIO_Sales_Data is missing expected column(s): "
            + ", ".join(missing)
        )


def prepare_pio_sales_for_eda(df):
    """Return an in-memory analysis copy with parsed dates and numeric measures."""
    validate_pio_columns(df)
    analysis = df.copy()

    invoice_text = (
        analysis["PIS_MST_IVC_DT"]
        .astype("string")
        .str.replace(r"\.0$", "", regex=True)
    )
    analysis[INVOICE_DATE] = pd.to_datetime(
        invoice_text, format="%Y%m%d", errors="coerce"
    )
    analysis[ALTERNATE_DATE] = pd.to_datetime(
        analysis["Deliminated Date"], errors="coerce"
    )

    month_text = (
        analysis["YYYYMM"].astype("string").str.replace(r"\.0$", "", regex=True)
    )
    analysis[MONTH] = pd.to_datetime(
        month_text, format="%Y%m", errors="coerce"
    )
    analysis[QUANTITY] = pd.to_numeric(
        analysis["SumOfPIS_INST_QT"], errors="coerce"
    )
    analysis[REVENUE] = pd.to_numeric(
        analysis["SumOfPIS_CRP_CFM_PRI"], errors="coerce"
    )
    analysis[UNIT_PRICE] = (
        analysis[REVENUE].div(analysis[QUANTITY])
        .where(analysis[QUANTITY] > 0)
    )
    return analysis


def missing_value_summary(df):
    """Summarize missing values in the original source columns."""
    summary = pd.DataFrame({
        "column": df.columns.astype(str),
        "missing_count": df.isna().sum().to_numpy(),
        "missing_percent": (df.isna().mean().to_numpy() * 100).round(2),
    })
    return summary.sort_values(
        ["missing_count", "column"], ascending=[False, True]
    ).reset_index(drop=True)


def overview_summary(raw_df, analysis_df):
    """Return high-level PIO coverage, totals, and cardinalities."""
    valid_invoice_dates = analysis_df[INVOICE_DATE].dropna()
    valid_months = analysis_df[MONTH].dropna()
    positive_quantity = analysis_df[QUANTITY] > 0
    priced_quantity = analysis_df.loc[positive_quantity, QUANTITY].sum(min_count=1)
    priced_revenue = analysis_df.loc[positive_quantity, REVENUE].sum(min_count=1)
    return {
        "rows": int(raw_df.shape[0]),
        "columns": int(raw_df.shape[1]),
        "duplicate_rows": int(raw_df.duplicated().sum()),
        "invoice_date_min": valid_invoice_dates.min() if not valid_invoice_dates.empty else None,
        "invoice_date_max": valid_invoice_dates.max() if not valid_invoice_dates.empty else None,
        "yyyymm_min": valid_months.min() if not valid_months.empty else None,
        "yyyymm_max": valid_months.max() if not valid_months.empty else None,
        "total_quantity": analysis_df[QUANTITY].sum(min_count=1),
        "total_revenue": analysis_df[REVENUE].sum(min_count=1),
        "average_unit_price": (
            priced_revenue / priced_quantity
            if priced_quantity > 0
            else None
        ),
        "unique_brands": int(raw_df["PIS_CMP_KND"].nunique(dropna=True)),
        "unique_models": int(raw_df["Model"].nunique(dropna=True)),
        "unique_model_years": int(raw_df["PIS_MDL_YY"].nunique(dropna=True)),
        "unique_part_numbers": int(raw_df["PIS_PNO"].nunique(dropna=True)),
    }


def monthly_summary(analysis_df):
    """Summarize monthly completeness, quantity, and revenue."""
    result = (
        analysis_df.dropna(subset=[MONTH])
        .groupby(MONTH, as_index=False)
        .agg(
            row_count=(MONTH, "size"),
            quantity=(QUANTITY, "sum"),
            revenue=(REVENUE, "sum"),
            unique_models=("Model", "nunique"),
            unique_parts=("PIS_PNO", "nunique"),
        )
        .sort_values(MONTH)
        .rename(columns={MONTH: "month"})
    )
    result["average_unit_price"] = (
        result["revenue"].div(result["quantity"]).where(result["quantity"] > 0)
    )
    return result


def model_part_sparsity_summary(analysis_df):
    """Summarize how many months each model-part combination is active."""
    result = (
        analysis_df.groupby(
            ["Model", "PIS_PNO", "Part Description"],
            dropna=False,
            as_index=False,
        )
        .agg(
            active_months=(MONTH, "nunique"),
            total_quantity=(QUANTITY, "sum"),
            total_revenue=(REVENUE, "sum"),
        )
    )
    result["avg_monthly_quantity"] = (
        result["total_quantity"]
        .div(result["active_months"])
        .where(result["active_months"] > 0)
    )
    return result.sort_values(
        ["active_months", "total_quantity"],
        ascending=[True, False],
    ).reset_index(drop=True)


def grouped_sales_summary(analysis_df, group_columns):
    """Aggregate quantity and revenue for one or more business dimensions."""
    result = (
        analysis_df.groupby(group_columns, dropna=False, as_index=False)
        .agg(quantity=(QUANTITY, "sum"), revenue=(REVENUE, "sum"))
    )
    result["average_unit_price"] = (
        result["revenue"].div(result["quantity"]).where(result["quantity"] > 0)
    )
    return result


def _top_n_share(summary, value_column, top_n):
    total = summary[value_column].sum()
    if pd.isna(total) or total == 0:
        return None
    return float(summary.nlargest(top_n, value_column)[value_column].sum() / total)


def concentration_summary(by_model, by_part, by_model_part, top_n=10):
    """Calculate top-N quantity and revenue shares for key product grains."""
    groups = {
        "model": by_model,
        "part": by_part,
        "model-part": by_model_part,
    }
    return pd.DataFrame([
        {
            "entity": entity,
            "top_n": top_n,
            "group_count": int(len(summary)),
            "top_n_quantity_share": _top_n_share(summary, "quantity", top_n),
            "top_n_revenue_share": _top_n_share(summary, "revenue", top_n),
        }
        for entity, summary in groups.items()
    ])


def build_eda_tables(analysis_df, top_n=10):
    """Create the requested read-only aggregation tables."""
    by_brand = grouped_sales_summary(analysis_df, ["PIS_CMP_KND"])
    by_model = grouped_sales_summary(analysis_df, ["Model"])
    by_part = grouped_sales_summary(
        analysis_df, ["PIS_PNO", "Part Description"]
    )
    by_model_part = grouped_sales_summary(
        analysis_df, ["Model", "PIS_PNO", "Part Description"]
    )
    return {
        "monthly": monthly_summary(analysis_df),
        "model_part_sparsity": model_part_sparsity_summary(analysis_df),
        "concentration": concentration_summary(
            by_model, by_part, by_model_part, top_n=top_n
        ),
        "by_brand": by_brand.sort_values("revenue", ascending=False),
        "top_models_by_quantity": by_model.nlargest(top_n, "quantity"),
        "top_models_by_revenue": by_model.nlargest(top_n, "revenue"),
        "top_parts_by_quantity": by_part.nlargest(top_n, "quantity"),
        "top_parts_by_revenue": by_part.nlargest(top_n, "revenue"),
        "top_model_parts_by_quantity": by_model_part.nlargest(top_n, "quantity"),
        "top_model_parts_by_revenue": by_model_part.nlargest(top_n, "revenue"),
    }


def quality_check_summary(raw_df, analysis_df, lower_quantile=0.005, upper_quantile=0.995):
    """Count basic quality conditions; unusual prices are heuristic flags."""
    invoice_nonempty = raw_df["PIS_MST_IVC_DT"].notna()
    alternate_nonempty = raw_df["Deliminated Date"].notna()
    month_nonempty = raw_df["YYYYMM"].notna()
    model_missing = (
        raw_df["Model"].isna()
        | raw_df["Model"].astype("string").str.strip().eq("")
    )
    part_missing = (
        raw_df["PIS_PNO"].isna()
        | raw_df["PIS_PNO"].astype("string").str.strip().eq("")
    )
    part_description_missing = (
        raw_df["Part Description"].isna()
        | raw_df["Part Description"].astype("string").str.strip().eq("")
    )

    positive_prices = analysis_df.loc[analysis_df[UNIT_PRICE] > 0, UNIT_PRICE]
    if positive_prices.empty:
        low_price = high_price = None
        unusual_price = pd.Series(False, index=analysis_df.index)
    else:
        low_price = float(positive_prices.quantile(lower_quantile))
        high_price = float(positive_prices.quantile(upper_quantile))
        unusual_price = (
            analysis_df[UNIT_PRICE].notna()
            & (
                (analysis_df[UNIT_PRICE] < low_price)
                | (analysis_df[UNIT_PRICE] > high_price)
            )
        )

    checks = [
        (
            "Invalid invoice date",
            invoice_nonempty & analysis_df[INVOICE_DATE].isna(),
            "Non-empty PIS_MST_IVC_DT did not parse as YYYYMMDD.",
        ),
        (
            "Invalid derived/normalized date",
            alternate_nonempty & analysis_df[ALTERNATE_DATE].isna(),
            "Non-empty Deliminated Date did not parse as a date; exact field meaning needs confirmation.",
        ),
        (
            "Invalid YYYYMM",
            month_nonempty & analysis_df[MONTH].isna(),
            "Non-empty YYYYMM did not parse as YYYYMM.",
        ),
        ("Missing model", model_missing, "Model is blank or missing."),
        ("Missing part number", part_missing, "PIS_PNO is blank or missing."),
        (
            "Missing part description",
            part_description_missing,
            "Part Description is blank or missing.",
        ),
        (
            "Negative quantity",
            analysis_df[QUANTITY] < 0,
            "Installed quantity is negative; confirm returns/reversals.",
        ),
        (
            "Negative revenue",
            analysis_df[REVENUE] < 0,
            "Revenue is negative; confirm credits/reversals.",
        ),
        (
            "Zero quantity with positive revenue",
            (analysis_df[QUANTITY] == 0) & (analysis_df[REVENUE] > 0),
            "Revenue exists with zero installed quantity.",
        ),
        (
            "Positive quantity with zero revenue",
            (analysis_df[QUANTITY] > 0) & (analysis_df[REVENUE] == 0),
            "Installed quantity exists with zero revenue.",
        ),
        (
            "Unusual unit price",
            unusual_price,
            (
                f"Unit price outside global {lower_quantile:.1%}–"
                f"{upper_quantile:.1%} bounds "
                f"({low_price!s} to {high_price!s}); diagnostic only."
            ),
        ),
    ]
    return pd.DataFrame([
        {
            "check": name,
            "row_count": int(mask.fillna(False).sum()),
            "details": details,
        }
        for name, mask, details in checks
    ])
