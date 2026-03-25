# Mouse Cage Card Generator

A small Shiny for Python app that converts a SoftMouse `.xlsx` export into a printable `notecards.xlsx` file for cage cards.

## What it does

- Upload a SoftMouse Excel sheet
- Enter cage-card settings in the app:
  - PI name
  - protocol number
  - contact name
  - contact phone
  - species
- Choose whether to include comments
- Download a formatted Excel file of printable cage cards

## Files

- `app.py` — Shiny app
- `notecard.py` — cage-card generation logic
- `requirements.txt` — Python dependencies

## Run locally

```bash
pip install -r requirements.txt
shiny run --reload app.py
```

## Run online

Visit posit connect cloud link on right hand side.
