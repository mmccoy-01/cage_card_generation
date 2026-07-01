# Mouse Cage Card Generator

A Shiny for Python app that converts SoftMouse `.xlsx` exports into print-ready PDF cage cards.

Inspired by [`mfe7/mouse_cage_cards`](https://github.com/mfe7/mouse_cage_cards).

Live app: [Posit Connect Cloud](https://019d274e-8c19-ccbd-b9f6-2762e561ccd9.share.connect.posit.cloud/)

## Features

- Upload a SoftMouse cage workbook `.xlsx`
- Optionally upload a SoftMouse mating workbook `.xlsx` to add sire/dam information on card backs
- Select or deselect cages using the `Print Card?` column
- Edit uploaded cage data directly in the app before printing
- Download a print-ready PDF
- Prints 4 cage cards per page
- Sorts mating cages first, then stock cages
- Adds mating cage back cards with sire, dam, genotype, litter line, and litter history
- Leaves stock cage backs blank for duplex-safe printing
- Normalizes dates to `YYYY-MM-DD`
- Supports sidebar defaults for PI, protocol, approval/expiration dates, contact info, species, and source

## Printing

The generated PDF is designed for duplex printing. Use these print settings:

```text
Scale: Actual Size / 100%
Double-sided: On
Flip: Long edge
````

## Files

* `app.py` — Shiny app interface and upload workflow
* `notecard.py` — PDF cage-card generation logic
* `requirements.txt` — Python dependencies

## Requirements

```txt
shiny
pandas
openpyxl
PyYAML
reportlab
xlsxwriter
```

## Run locally

```bash
pip install -r requirements.txt
shiny run --reload app.py
```
