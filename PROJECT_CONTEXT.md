# PROJECT_CONTEXT.md

## Hyundai Mobis Capstone Summary

This project is a UCLA MEng Data Science capstone project with Hyundai Mobis / Mobis Parts America.

The project name is:

**PIO Accessories Forecasting Optimization**

PIO means **Port-Installed Options**. These are accessories installed at the port before vehicles are delivered to dealers or customers.

The project focuses on **PIO accessories**, not service parts or maintenance parts.

## Business Understanding From Sponsor Meetings

The sponsor explained that PIO accessories are important for demand planning, inventory planning, monthly sales planning, and month-end business reporting.

The main business goal is to improve forecast accuracy for PIO accessory demand and sales dollars.

The sponsor wants to reduce the forecast gap. For example, instead of a broad sales forecast range such as **$28M-$30M**, the goal is to estimate more accurately whether actual month-end PIO sales will land closer to **$28M, $28.5M, $29M**, etc.

The project should consider both:

* total monthly PIO revenue / sales dollars
* model-by-accessory quantity and revenue

Quantity forecasting is important for demand and inventory planning. Revenue forecasting is important for monthly sales planning and business reporting.

## Current Data Understanding

The current Excel workbook includes at least two important sheets:

* `PIO_Sales_Data`
* `Vehicle_Wholesale_Data`

`PIO_Sales_Data` appears to be relatively clean and long-format. It includes PIO sales information such as invoice date, brand/company code, model or series code, model year, part number, installed quantity, sales dollars, month, model, and part description.

`Vehicle_Wholesale_Data` appears to be more like a formatted Excel report. It may require additional cleaning because Wholesale and Fleet information may be embedded in the same sheet.

The raw workbook should not be assumed to be clean.

## Important Business Concepts

Relevant concepts include:

* PIO
* vehicle wholesale volume
* fleet volume
* total vehicle volume
* penetration rate
* model
* model year
* model code
* part number
* part description
* installed quantity
* revenue / sales dollars
* inventory
* AMD
* MOS
* working days
* PMV / PMVW or similar per-vehicle wholesale accessory dollar metrics

Some acronyms and business definitions are still being confirmed with Mobis, so assumptions should be documented clearly.

## Forecasting Direction

A useful business-driven forecasting idea is:

```text
forecast PIO quantity = forecast vehicle volume × predicted penetration rate
```

Revenue may then be estimated using quantity and price:

```text
forecast PIO revenue = forecast PIO quantity × price
```

Forecasts can be built at the model-accessory level and then aggregated to model, brand, and total PIO dollars.

This approach should be tested against simple baselines such as naive forecast, moving average, and historical average.

The project should not start with complex neural networks. Regression-based models or XGBoost may be considered later if the data supports them.

## Recommended Workflow

1. Inspect the raw Excel workbook.
2. Understand each sheet and column.
3. Clean PIO sales data.
4. Clean vehicle wholesale and fleet data.
5. Create monthly analysis tables.
6. Merge PIO sales with vehicle volume where possible.
7. Calculate penetration rate only when the denominator is valid and clearly matched.
8. Perform EDA.
9. Build baseline forecasts.
10. Evaluate forecasts using WAPE, RMSE, Bias, over-forecast rate, and under-forecast rate.
11. Track month-end landing forecast accuracy at different points in the month.
12. Optionally connect results to a dashboard later.

## Important Caution

The dashboard should be treated as a visualization layer.

The core analytical deliverable should be the data cleaning logic, forecasting method, documented assumptions, and model evaluation.

Daily or monthly sales may be affected by working days, port schedules, weekend data rollups, price changes, and other operational factors. These should be considered during EDA and forecasting.
