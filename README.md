# AP Excel Validator

A Streamlit app that validates filled Academic Personnel Excel files against the template rules.

## Features

- Upload `.xlsx` files and validate against 24 column rules (required fields, dropdowns, dates, numbers)
- Select which columns to validate with an interactive checkbox grid
- Summary dashboard with metrics and error-per-column bar chart
- Color-coded data table (green = valid, red = error, grey = not validated)
- Expandable error details per row with filtering

## Install

```bash
pip install -r requirements.txt
```

## Run

```bash
streamlit run validator_app.py
```

Then upload a filled `AcademicPersonel_update.xlsx` file through the sidebar.
