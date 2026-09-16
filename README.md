# FP&A Performance Dashboard

Automated variance analysis, forecasting, and executive reporting for a multi-business-unit finance team.

## The Problem

Month-end close across multiple business units usually means the same manual routine: pull actuals from each unit, reconcile against budget, calculate variances by hand, flag what's favourable or unfavourable, roll the numbers into a forecast, and write commentary explaining what happened. It's repetitive, error-prone under deadline, and eats time that should go into actually interpreting the numbers.

This project automates that routine end to end, from raw consolidation through to an executive-ready dashboard with AI-assisted commentary.

## What It Does

- Consolidates results for Business Units A, B, and C
- Calculates YTD Actual vs Budget variance, with favourable/unfavourable flags
- Annualizes period budgets (YTD Budget × 3) for full-year comparability
- Builds a rolling full-year forecast for the remaining months
- Exports dashboard-ready Excel files
- Feeds a Power BI executive dashboard with:
  - KPI cards with period-over-period comparisons
  - Business unit performance chart
  - Monthly revenue and net income trend
  - YTD variance analysis table
  - Management commentary panel
- Generates executive commentary automatically via n8n, connected directly to Power BI

## Tech Stack

- **Python** — consolidation, variance calculation, forecasting, Excel export
- **Power BI** — executive dashboard and visualization layer
- **n8n** — automated commentary generation, connected directly to the dashboard

## How It Works

1. Business unit results are consolidated into a single dataset
2. YTD Actual vs Budget variance is calculated per unit, with direction flagged
3. Period budgets are annualized (YTD Budget × 3) for full-year comparison against forecast
4. Remaining months are projected with a rolling forecast method
5. Outputs are exported to Excel in a dashboard-ready format
6. Power BI ingests the output and renders the dashboard
7. n8n generates commentary on the variance results and pushes it into the dashboard's commentary panel

## Data Note

This repository does not include the original input dataset, excluded for licensing reasons. The code is written to run against any similarly structured multi-business-unit actuals/budget dataset — swap in your own to use it.

## Roadmap

- Chart of accounts mapping for multi-source consolidation
- Scenario-based forecasting (best case / base case / worst case)
- Driver-based forecast inputs instead of straight-line projection

## Background

This project started as an extension of a CFI (Corporate Finance Institute) FP&A challenge. The original was a manual Excel exercise; this version automates the same workflow in Python, adds a rolling forecast and AI-assisted commentary, and delivers it as a live Power BI dashboard instead of a static spreadsheet.

## License

MIT
