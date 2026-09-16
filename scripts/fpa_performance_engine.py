"""
FP&A Performance Engine

Consolidates multi-unit financial results, calculates YTD variance analysis,
builds a rolling full-year forecast, and exports dashboard-ready outputs.
"""

import pandas as pd
import numpy as np
import os
import json

# Paths
base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
file_path = os.path.join(base_dir, "input", "CFI_Challenge.xlsx")
output_dir = os.path.join(base_dir, "output")
os.makedirs(output_dir, exist_ok=True)


def load_unit(sheet_name: str) -> pd.DataFrame:
    """Load income statement and balance sheet lines for one business unit."""
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

    income_df = pd.DataFrame({
        "Line Item": ["Revenue", "Cost of Goods Sold", "Gross Profit", "SG&A Expenses", "Net Income"],
        "Apr": df.iloc[8:13, 4].values,
        "May": df.iloc[8:13, 5].values,
        "Jun": df.iloc[8:13, 6].values,
        "Jul": df.iloc[8:13, 7].values,
        "YTD Budget": df.iloc[8:13, 10].values,
    })
    income_df["YTD Actual"] = income_df[["Apr", "May", "Jun", "Jul"]].sum(axis=1)
    income_df["Type"] = "Income Statement"

    balance_df = pd.DataFrame({
        "Line Item": ["Cash Balance", "Accounts Receivable", "Accounts Payable"],
        "Apr": df.iloc[15:18, 4].values,
        "May": df.iloc[15:18, 5].values,
        "Jun": df.iloc[15:18, 6].values,
        "Jul": df.iloc[15:18, 7].values,
        "YTD Budget": df.iloc[15:18, 10].values,
    })
    balance_df["YTD Actual"] = balance_df["Jul"]
    balance_df["Type"] = "Balance Sheet"

    return pd.concat([income_df, balance_df], ignore_index=True)


def flag_variance(row):
    """Basic favourable / unfavourable flag."""
    cost_like = row["Line Item"] in ["Cost of Goods Sold", "SG&A Expenses", "Accounts Payable"]
    if cost_like:
        return "Favourable" if row["Variance $"] > 0 else "Unfavourable"
    return "Favourable" if row["Variance $"] > 0 else "Unfavourable"


def create_forecast(unit_df: pd.DataFrame) -> pd.DataFrame:
    """
    Forecast remaining months using recent average and conservative growth.
    Full Year Budget is annualized as YTD Budget x 3.
    """
    forecast_df = unit_df[unit_df["Type"] == "Income Statement"].copy()

    actual_months = ["Apr", "May", "Jun", "Jul"]
    forecast_months = ["Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]

    forecast_df["Avg Monthly"] = forecast_df[actual_months].mean(axis=1)

    growth_list = []
    for i in range(len(actual_months) - 1):
        prev_month = forecast_df[actual_months[i]]
        curr_month = forecast_df[actual_months[i + 1]]
        growth = (curr_month - prev_month) / prev_month.replace(0, np.nan).abs()
        growth_list.append(growth)

    avg_growth = pd.concat(growth_list, axis=1).mean(axis=1).fillna(0)
    conservative_growth = avg_growth * 0.5

    last_value = forecast_df["Jul"]
    for i, month in enumerate(forecast_months):
        if i == 0:
            forecast_df[month] = last_value * (1 + conservative_growth)
        else:
            prev_month = forecast_months[i - 1]
            forecast_df[month] = forecast_df[prev_month] * (1 + conservative_growth)

    forecast_df["Full Year Forecast"] = (
        forecast_df[actual_months].sum(axis=1) + forecast_df[forecast_months].sum(axis=1)
    )
    forecast_df["Full Year Budget"] = forecast_df["YTD Budget"] * 3
    forecast_df["FY Variance $"] = forecast_df["Full Year Forecast"] - forecast_df["Full Year Budget"]
    forecast_df["FY Variance %"] = (
        forecast_df["FY Variance $"] / forecast_df["Full Year Budget"].replace(0, np.nan) * 100
    ).round(1)

    cols = (
        ["Line Item"]
        + actual_months
        + forecast_months
        + ["Full Year Forecast", "Full Year Budget", "FY Variance $", "FY Variance %", "YTD Budget"]
    )
    return forecast_df[cols]


# Load units
df_a = load_unit("Business A")
df_b = load_unit("Business B")
df_c = load_unit("Business C")

print("Business A")
print(df_a.to_string(index=False))
print("\nBusiness B")
print(df_b.to_string(index=False))
print("\nBusiness C")
print(df_c.to_string(index=False))

# Group YTD consolidation
group = df_a[["Line Item", "Type"]].copy()
group["YTD Actual"] = df_a["YTD Actual"] + df_b["YTD Actual"] + df_c["YTD Actual"]
group["YTD Budget"] = df_a["YTD Budget"] + df_b["YTD Budget"] + df_c["YTD Budget"]
group["Variance $"] = group["YTD Actual"] - group["YTD Budget"]
group["Variance %"] = (group["Variance $"] / group["YTD Budget"].replace(0, np.nan) * 100).round(1)
group["F/U"] = group.apply(flag_variance, axis=1)

print("\nGroup consolidated view")
print(group.to_string(index=False))

# Unit forecasts
forecast_a = create_forecast(df_a)
forecast_b = create_forecast(df_b)
forecast_c = create_forecast(df_c)

print("\nForecast - Business A")
print(forecast_a.to_string(index=False))
print("\nForecast - Business B")
print(forecast_b.to_string(index=False))
print("\nForecast - Business C")
print(forecast_c.to_string(index=False))

# Group forecast
forecast_cols = [
    "Line Item", "Apr", "May", "Jun", "Jul",
    "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar",
    "Full Year Forecast", "Full Year Budget", "FY Variance $", "FY Variance %", "YTD Budget"
]

fa = forecast_a[forecast_cols].copy()
fb = forecast_b[forecast_cols].copy()
fc = forecast_c[forecast_cols].copy()

group_forecast = fa.copy()
numeric_cols = [
    "Apr", "May", "Jun", "Jul",
    "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar",
    "Full Year Forecast", "Full Year Budget", "FY Variance $", "YTD Budget"
]

for col in numeric_cols:
    group_forecast[col] = fa[col] + fb[col] + fc[col]

group_forecast["FY Variance %"] = (
    group_forecast["FY Variance $"] / group_forecast["Full Year Budget"].replace(0, np.nan) * 100
).round(1)

group_forecast["FY F/U"] = group_forecast.apply(
    lambda row: "Favourable" if row["FY Variance $"] > 0 else "Unfavourable",
    axis=1
)

print("\nGroup forecast")
print(group_forecast.to_string(index=False))

# Save core outputs
group.to_excel(os.path.join(output_dir, "Group_Variance.xlsx"), index=False)
group_forecast.to_excel(os.path.join(output_dir, "Group_Forecast.xlsx"), index=False)
forecast_a.to_excel(os.path.join(output_dir, "Forecast_Unit_A.xlsx"), index=False)
forecast_b.to_excel(os.path.join(output_dir, "Forecast_Unit_B.xlsx"), index=False)
forecast_c.to_excel(os.path.join(output_dir, "Forecast_Unit_C.xlsx"), index=False)

# KPI summary
chart_summary = pd.DataFrame([
    {
        "Metric": "YTD Revenue",
        "Actual": int(group.loc[group["Line Item"] == "Revenue", "YTD Actual"].values[0]),
        "Budget": int(group.loc[group["Line Item"] == "Revenue", "YTD Budget"].values[0]),
        "Variance": int(group.loc[group["Line Item"] == "Revenue", "Variance $"].values[0]),
    },
    {
        "Metric": "YTD Net Income",
        "Actual": int(group.loc[group["Line Item"] == "Net Income", "YTD Actual"].values[0]),
        "Budget": int(group.loc[group["Line Item"] == "Net Income", "YTD Budget"].values[0]),
        "Variance": int(group.loc[group["Line Item"] == "Net Income", "Variance $"].values[0]),
    },
    {
        "Metric": "FY Revenue",
        "Actual": int(group_forecast.loc[group_forecast["Line Item"] == "Revenue", "Full Year Forecast"].values[0]),
        "Budget": int(group_forecast.loc[group_forecast["Line Item"] == "Revenue", "Full Year Budget"].values[0]),
        "Variance": int(group_forecast.loc[group_forecast["Line Item"] == "Revenue", "FY Variance $"].values[0]),
    },
    {
        "Metric": "FY Net Income",
        "Actual": int(group_forecast.loc[group_forecast["Line Item"] == "Net Income", "Full Year Forecast"].values[0]),
        "Budget": int(group_forecast.loc[group_forecast["Line Item"] == "Net Income", "Full Year Budget"].values[0]),
        "Variance": int(group_forecast.loc[group_forecast["Line Item"] == "Net Income", "FY Variance $"].values[0]),
    },
])

unit_chart = pd.DataFrame([
    {
        "Unit": "Unit A",
        "FY_NI_Forecast": int(forecast_a.loc[forecast_a["Line Item"] == "Net Income", "Full Year Forecast"].values[0]),
        "FY_NI_Budget": int(forecast_a.loc[forecast_a["Line Item"] == "Net Income", "Full Year Budget"].values[0]),
    },
    {
        "Unit": "Unit B",
        "FY_NI_Forecast": int(forecast_b.loc[forecast_b["Line Item"] == "Net Income", "Full Year Forecast"].values[0]),
        "FY_NI_Budget": int(forecast_b.loc[forecast_b["Line Item"] == "Net Income", "Full Year Budget"].values[0]),
    },
    {
        "Unit": "Unit C",
        "FY_NI_Forecast": int(forecast_c.loc[forecast_c["Line Item"] == "Net Income", "Full Year Forecast"].values[0]),
        "FY_NI_Budget": int(forecast_c.loc[forecast_c["Line Item"] == "Net Income", "Full Year Budget"].values[0]),
    },
])

chart_summary.to_excel(os.path.join(output_dir, "Chart_KPI_Summary.xlsx"), index=False)
unit_chart.to_excel(os.path.join(output_dir, "Chart_Unit_Performance.xlsx"), index=False)

# Executive metrics
executive_metrics = {
    "YTD_Revenue_Actual": int(group.loc[group["Line Item"] == "Revenue", "YTD Actual"].values[0]),
    "YTD_Revenue_Variance": int(group.loc[group["Line Item"] == "Revenue", "Variance $"].values[0]),
    "YTD_Net_Income_Actual": int(group.loc[group["Line Item"] == "Net Income", "YTD Actual"].values[0]),
    "YTD_Net_Income_Variance": int(group.loc[group["Line Item"] == "Net Income", "Variance $"].values[0]),
    "YTD_Gross_Margin_Pct": round(
        (group.loc[group["Line Item"] == "Gross Profit", "YTD Actual"].values[0] /
         group.loc[group["Line Item"] == "Revenue", "YTD Actual"].values[0]) * 100, 1
    ),
    "YTD_Net_Margin_Pct": round(
        (group.loc[group["Line Item"] == "Net Income", "YTD Actual"].values[0] /
         group.loc[group["Line Item"] == "Revenue", "YTD Actual"].values[0]) * 100, 1
    ),
    "YTD_COGS_Variance": int(group.loc[group["Line Item"] == "Cost of Goods Sold", "Variance $"].values[0]),
    "YTD_SGA_Variance": int(group.loc[group["Line Item"] == "SG&A Expenses", "Variance $"].values[0]),
    "Unit_A_Net_Income": int(df_a.loc[df_a["Line Item"] == "Net Income", "YTD Actual"].values[0]),
    "Unit_B_Net_Income": int(df_b.loc[df_b["Line Item"] == "Net Income", "YTD Actual"].values[0]),
    "Unit_C_Net_Income": int(df_c.loc[df_c["Line Item"] == "Net Income", "YTD Actual"].values[0]),
    "FY_Revenue_Forecast": int(group_forecast.loc[group_forecast["Line Item"] == "Revenue", "Full Year Forecast"].values[0]),
    "FY_Revenue_Budget": int(group_forecast.loc[group_forecast["Line Item"] == "Revenue", "Full Year Budget"].values[0]),
    "FY_Net_Income_Forecast": int(group_forecast.loc[group_forecast["Line Item"] == "Net Income", "Full Year Forecast"].values[0]),
    "FY_Net_Income_Budget": int(group_forecast.loc[group_forecast["Line Item"] == "Net Income", "Full Year Budget"].values[0]),
}

with open(os.path.join(output_dir, "Executive_Metrics.json"), "w") as f:
    json.dump(executive_metrics, f, indent=4)

print("\nFiles saved to output folder")
print("Run complete")