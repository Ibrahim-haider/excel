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
