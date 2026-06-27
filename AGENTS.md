# AGENTS.md

## Project Name

PIO Accessories Forecasting Optimization

## Project Context

This is a UCLA MEng Data Science capstone project with Hyundai Mobis / Mobis Parts America.

The project focuses on PIO accessory demand and sales forecasting.

PIO means Port-Installed Options: accessories installed at the port before vehicles are delivered to dealers or customers.

The project focus is PIO/accessory forecasting, not service parts or maintenance parts.

## Business Goal

The goal is to improve PIO accessory demand and sales forecasting by making forecasts closer to actual results.

The sponsor wants to reduce the forecast gap and improve month-end forecast accuracy. For example, instead of a broad sales forecast range such as $28M-$30M, the goal is to better estimate whether actual month-end PIO sales will land closer to $28M, $28.5M, $29M, etc.

The project should focus on both:

* PIO installation quantity / demand forecasting
* PIO revenue / sales dollar forecasting

Quantity forecasting is important for demand and inventory planning. Revenue forecasting is important for monthly sales planning and business reporting.

The forecasting logic may involve:

* historical PIO sales
* installation quantity
* vehicle model
* model year
* part number
* part description
* revenue / sales dollars
* wholesale vehicle volume
* fleet vehicle volume
* total vehicle volume
* penetration rate
* inventory
* AMD
* MOS
* working days
* PMV / PMVW or similar per-vehicle wholesale accessory dollar metrics

Some acronym definitions are still being confirmed, so do not hardcode business assumptions without documenting them.

## Current Data

The current available dataset is a sales-related Excel workbook.

Known sheets include:

* `PIO_Sales_Data`
* `Vehicle_Wholesale_Data`

`PIO_Sales_Data` appears to be a relatively clean long-format table with fields such as invoice date, company/brand code, model/series code, model year, part number, installed quantity, sales dollars, YYYYMM, model, and part description.

`Vehicle_Wholesale_Data` is not a clean table. It is formatted like an Excel report and includes both Wholesale and Fleet sections in the same sheet. It may contain:

* formatted Excel tables
* title rows
* merged cells
* frozen rows
* multi-row headers
* year headers such as 2025 and 2026
* repeated month columns such as Jan, Feb, Mar
* wide-format month columns
* subtotal rows such as HMA Total, GMA Total, H+G Total, KUS Total, and Grand Total
* missing values caused by merged cells
* unclear or inconsistent field names

Do not assume the raw Excel is clean.

## Current Goal

Build a Python EDA and data cleaning pipeline first.

Do not build a complicated forecasting model until the data structure is understood.

First deliverables:

1. inspect workbook sheet names and structure
2. identify useful columns and business meaning
3. clean the PIO sales data into a standard monthly table
4. clean the vehicle sales data into a standard monthly table
5. merge PIO sales with vehicle volume where possible
6. calculate penetration rate where possible
7. perform EDA
8. prepare the data for later forecasting

## Preferred Clean Data Format

### PIO Monthly Table

Try to transform PIO sales data into this long-format structure:

month | brand | model | model_code | model_year | part_number | part_description | quantity | revenue

If some fields are unavailable, keep the available fields and document assumptions.

### Vehicle Monthly Table

Try to transform vehicle sales data into this long-format structure:

month | channel | brand | model | model_code | vehicle_volume

where channel may include:

* Wholesale
* Fleet
* Total, if it is directly provided or can be safely calculated

### Merged Analysis Table

When possible, create a merged model-part-month table:

month | brand | model | model_code | model_year | part_number | part_description | quantity | revenue | wholesale_volume | fleet_volume | total_vehicle_volume | fleet_share | penetration_rate

Penetration rate can be calculated as:

penetration_rate = PIO quantity / matched vehicle volume

Only calculate this when the vehicle volume denominator is valid and clearly matched. Depending on the business definition and data match, the denominator may be wholesale volume, fleet volume, or total vehicle volume. Do not assume total vehicle volume is always the correct denominator.

PMV / PMVW can be treated as a per-vehicle accessory dollar metric, but the exact acronym definition should be confirmed with Mobis.

A general calculation is:

PMV or PMVW = PIO revenue / matched vehicle wholesale volume

Only calculate this when the denominator is valid and clearly matched.

## Data Cleaning Notes

Be careful when merging PIO sales data with vehicle sales data.

Model code alone may not be unique because multiple model variants can share the same model code. For example, a gasoline model, HEV model, and PHEV model may use the same model code.

Prefer using a combination of:

* month
* brand
* model name
* model code, if reliable

Document any manual model-name mapping.

Brand codes may need mapping. For example, PIO data may use codes such as H, K, or G, while vehicle data may use HMA, KUS, or GMA. Confirm definitions before hardcoding assumptions.

Daily or monthly sales may be affected by working days, port schedules, weekend data rollups, and price changes. Use working-day normalization or rolling averages where appropriate.

## Modeling Direction

Start simple.

Preferred sequence:

1. naive baseline
2. moving average baseline
3. historical penetration rate × vehicle volume baseline
4. seasonal baseline only if monthly seasonality exists and enough history is available
5. regression-based driver model
6. XGBoost / tree-based model only if enough structured features exist

Avoid neural networks unless there is a strong reason.

A useful business-driven forecasting logic is:

forecast PIO quantity = forecast vehicle volume × predicted penetration rate

Revenue forecast can then be calculated as:

forecast PIO revenue = forecast PIO quantity × price

Forecasts can be aggregated from model-part level to model, brand, and total PIO dollars.

This should be tested against simpler baselines.

## Evaluation Metrics

Use:

* WAPE
* RMSE
* Bias
* over-forecast rate
* under-forecast rate

MAPE can be used only when actual values are not zero or near zero.

For sparse model-part combinations, MAPE may be misleading. Prefer WAPE and bias-based metrics.

For month-end landing forecasts, also track forecast error at different points in the month, such as mid-month, 10 working days before month-end, and 5 working days before month-end.

## Coding Style

Use simple, readable Python.
Use pandas, numpy, matplotlib, scikit-learn, and openpyxl.
Keep notebooks understandable.
Create reusable functions in src/.
Add comments explaining business assumptions.
Do not over-engineer the project.

## Data Safety

Do not commit raw company data.
Raw Excel files should stay under data/raw/.
Processed files should stay under data/processed/ or outputs/.
These folders are ignored by Git.
