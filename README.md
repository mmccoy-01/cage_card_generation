# Mouse Cage Card Generator

A small Shiny for Python app that converts a SoftMouse `.xlsx` export into a printable `notecards.xlsx` file for cage cards.

To run online: visit [posit connect cloud link](https://019d274e-8c19-ccbd-b9f6-2762e561ccd9.share.connect.posit.cloud/)

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
