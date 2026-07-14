# Executive Sales Intelligence Dashboard

This Streamlit application recreates and automates the company's manually prepared executive summary.

## Included metrics

- MTD Sales, Target and Achievement
- Cash Sales, Target and Achievement
- YTD Sales, Target and Achievement
- Days Remaining
- Required Run Rate
- Current Run Rate
- Risk Level
- Zone Performance
- Branch Categorization
- Road to Target
- MTD vs YTD Achievement
- Highest Value Contributors
- Daily Tracking
- Management Insights
- Excel and PDF exports

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy

Upload the project to GitHub and deploy `app.py` on Streamlit Community Cloud.


## Single-file workflow

This version requires only **one Excel workbook**.

The workbook must contain:
- `Working Sheet`
- `Raw Data`

The included demo workbook is `data/Ibrahim.xlsx`. You do not need separate MTD and YTD files.


## Fix included

Excel export now safely handles blank, NaN, and infinite calculated values, preventing
the `Cannot convert {0!r} to Excel` error on Streamlit Cloud.


## Flexible column mapping

The dashboard now recognizes common alternative column names and provides a manual mapping screen for unfamiliar names. The file must still contain equivalent sales concepts such as branch, target, MTD sales, YTD sales and cash sales.


## Excel-style column filters

The dashboard now includes per-column filters for the main tables:

- Multiselect filters for categories such as zone, branch, store and performance band
- Numeric range sliders for sales, targets and achievement values
- Date range filters for date columns
- Text search for high-cardinality columns
- Independent filter panels for zone performance, branch details, contributors and daily tracking
