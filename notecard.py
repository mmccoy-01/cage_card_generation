import argparse
import io
import re
import textwrap
import warnings
from datetime import date, datetime
from pathlib import Path
from typing import Any, BinaryIO

import pandas as pd
import xlsxwriter
import yaml
from openpyxl import load_workbook


MAX_MICE_PER_CAGE = 6
VISIBLE_COLS_PER_CARD = 6
GUTTER_COLS = 1
CARD_ROWS = 14
ROW_GAP = 2
ROWS_PER_SHEET = CARD_ROWS * 2 + ROW_GAP
RIGHT_CARD_START = VISIBLE_COLS_PER_CARD + GUTTER_COLS
PRINT_LAST_COL = RIGHT_CARD_START + VISIBLE_COLS_PER_CARD - 1

# Physical card target: 12.7 cm x 7.7 cm, i.e. 5.0 in x ~3.03 in.
CARD_WIDTH_PX = 480
CARD_ROW_HEIGHTS_PT = [19, 14, 14, 14, 17, 18, 16.3, 15, 15, 15, 15, 15, 15, 15]
ROW_GAP_HEIGHT_PT = 14

HEADER_NAMES = {
    "print_card": ["print card?", "print card", "print", "include", "include card", "selected"],
    "cage_tag": ["cage tag"],
    "num_mice": ["# of mice", "num mice", "number of mice"],
    # Optional backward compatibility only. Current status is inferred from Mating SID / Litter Mouseline.
    "disposition": ["disposition"],
    "mating_sid": ["mating sid", "mating id", "mating", "sid"],
    "litter_mouseline": [
        "litter mouseline",
        "litter mouse line",
        "litter strain",
        "litter line",
    ],
    "mice_tags": ["mice tags [sex, dob, age]", "mice tags", "mouse tags"],
    "genotypes": ["genotypes", "genotype"],
    "comment": ["comment", "comments", "notes"],
    "end_date": ["end date", "setup date"],
    "source": ["source"],
    "protocol_num": ["protocol", "protocol number", "protocol #", "protocol num"],
    "approved_date": ["approved", "approved date", "approval date"],
    "expires_date": ["expires", "expires date", "expiration date", "expiry date"],
}

MATING_HEADER_NAMES = {
    "cage_tag": ["cage tag"],
    "comment": ["comment", "comments", "notes"],
    "mating_sid": ["mating sid", "mating id", "mating", "sid"],
    "litter_mouseline": ["litter mouseline", "litter mouse line", "litter strain", "litter line"],
    "setup_date": ["set up date", "setup date"],
    "since_last_litter": ["since last litter"],
    "since_setup": ["since set up", "since setup"],
    "sire_genotype": ["m genotype", "male genotype", "sire genotype"],
    "dam_genotype": ["f genotype", "female genotype", "dam genotype"],
    "sire_tag": ["m tag", "male tag", "sire tag"],
    "female_mate_date": ["f mate date", "female mate date", "dam mate date"],
    "dam_tag": ["f tag", "female tag", "dam tag"],
    "num_litters": ["# litters", "num litters", "number of litters"],
    "litter_sids": ["litter sids [size, dob, wean date, state, end date, end reason]", "litter sids", "litter sid"],
}

DEFAULT_SETTINGS = {
    "PI_name": "",
    "protocol_num": "",
    "approved_date": "",
    "expires_date": "",
    "contact_name": "",
    "contact_phone": "",
    "species": "Mouse",
    "source": "",
}

FALSEY_PRINT_VALUES = {"false", "f", "no", "n", "0", "skip", "exclude", "dont print", "don't print", "unchecked"}
TRUEY_PRINT_VALUES = {"true", "t", "yes", "y", "1", "print", "include", "checked"}


def safe_str(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()




def normalize_identifier(value: Any) -> str:
    """Normalize IDs that Excel/pandas may coerce to numeric values, e.g. 176.0 -> 176."""
    if value is None:
        return ""
    if isinstance(value, float) and not pd.isna(value) and value.is_integer():
        return str(int(value))
    text = safe_str(value)
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


def normalize_settings(settings: dict[str, Any] | None) -> dict[str, str]:
    normalized = dict(DEFAULT_SETTINGS)
    if settings:
        normalized.update({k: safe_str(v) for k, v in settings.items()})
    if not normalized["species"]:
        normalized["species"] = "Mouse"
    return normalized


def normalize_header(value: Any) -> str:
    text = safe_str(value).lower()
    text = text.replace("\n", " ").replace("\r", " ").replace("_", " ").replace(".", " ")
    text = re.sub(r"^\$+\s*", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def build_header_index(header_row: list[Any], header_names: dict[str, list[str]] | None = None) -> dict[str, int | None]:
    names = header_names or HEADER_NAMES
    normalized = {normalize_header(v): i for i, v in enumerate(header_row) if safe_str(v)}
    out: dict[str, int | None] = {}
    for key, candidates in names.items():
        out[key] = None
        for name in candidates:
            if normalize_header(name) in normalized:
                out[key] = normalized[normalize_header(name)]
                break
    return out


def cell(data_row: list[Any], header_index: dict[str, int | None], key: str, default: Any = "") -> Any:
    idx = header_index.get(key)
    if idx is None or idx >= len(data_row):
        return default
    value = data_row[idx]
    return default if value is None else value


def cleaned_lines(value: Any, keep_blank_lines: bool = False) -> list[str]:
    text = safe_str(value)
    if not text:
        return [""] if keep_blank_lines else []
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if keep_blank_lines:
        return [line.strip() for line in lines]
    return [line.strip() for line in lines if line.strip()]


def safe_int(value: Any, default: int = 0) -> int:
    text = safe_str(value)
    if not text:
        return default
    match = re.search(r"\d+", text)
    return int(match.group(0)) if match else default


def should_print_card(value: Any) -> bool:
    """Treat a missing/blank print column as selected by default."""
    if value is None:
        return True
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if pd.isna(value):
            return True
        return value != 0
    text = safe_str(value).lower()
    if not text:
        return True
    if text in FALSEY_PRINT_VALUES:
        return False
    if text in TRUEY_PRINT_VALUES:
        return True
    return True


def format_date_value(value: Any) -> str:
    """Return a single date-like value as YYYY-MM-DD when possible."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")

    text = safe_str(value)
    if not text:
        return ""

    # Cells may arrive as strings like "2026-04-17 00:00:00".
    text = re.sub(r"\s+00:00:00$", "", text)

    for fmt in ("%Y-%m-%d", "%m-%d-%Y", "%m/%d/%Y", "%m-%d-%y", "%m/%d/%y"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass

    parsed = pd.to_datetime(text, errors="coerce")
    if not pd.isna(parsed):
        return parsed.strftime("%Y-%m-%d")

    return text


def normalize_dates_in_text(value: Any) -> str:
    """Normalize every visible date in a free-text/multiline field to YYYY-MM-DD."""
    text = safe_str(value)
    if not text:
        return ""

    def repl(match: re.Match[str]) -> str:
        return format_date_value(match.group(0))

    # Normalize mm-dd-yyyy and mm/dd/yyyy dates that appear inside SoftMouse multiline fields.
    text = re.sub(r"\b[0-1]?\d[-/][0-3]?\d[-/](?:20)?\d{2}\b", repl, text)
    text = re.sub(r"\b20\d{2}-[0-1]\d-[0-3]\d\b", repl, text)
    return text


def compact_multiline(value: Any, separator: str = "\n") -> str:
    lines = [normalize_dates_in_text(line) for line in cleaned_lines(value)]
    return separator.join(lines)


def parse_mouse_lines(mouse_lines: list[str]) -> list[dict[str, str]]:
    parsed = []
    for raw in mouse_lines:
        tag = raw.split("[")[0].strip()
        sex_match = re.search(r"\[(M|F)", raw, flags=re.IGNORECASE)
        dob_match = re.search(r"([0-1]?[0-9][-/][0-3]?[0-9][-/](?:20)?[0-9]{2}|20[0-9]{2}-[0-1][0-9]-[0-3][0-9])", raw)
        parsed.append(
            {
                "tag": tag,
                "sex": sex_match.group(1).upper() if sex_match else "",
                "dob": format_date_value(dob_match.group(1)) if dob_match else "",
                "raw": normalize_dates_in_text(raw),
            }
        )
    return parsed


def compact_note(comment: str, overflow_count: int = 0) -> str:
    parts = []
    note = " ".join(cleaned_lines(normalize_dates_in_text(comment)))
    if note:
        parts.append(note)
    if overflow_count > 0:
        parts.append(f"+{overflow_count} more mouse(s) not shown")
    if not parts:
        return ""
    return textwrap.shorten(" | ".join(parts), width=88, placeholder="...")


def set_column_widths(worksheet: xlsxwriter.worksheet.Worksheet) -> None:
    # Column pixels sum to 480 px for each 6-column card = 12.7 cm / 5.0 in at 96 px/in.
    card_col_widths_px = [64, 78, 64, 72, 78, 124]
    for i, width in enumerate(card_col_widths_px):
        worksheet.set_column_pixels(i, i, width)
    worksheet.set_column_pixels(VISIBLE_COLS_PER_CARD, VISIBLE_COLS_PER_CARD, 29)
    for i, width in enumerate(card_col_widths_px, start=RIGHT_CARD_START):
        worksheet.set_column_pixels(i, i, width)


def set_page_row_heights(worksheet: xlsxwriter.worksheet.Worksheet, page_top: int) -> None:
    for card_top in (page_top, page_top + CARD_ROWS + ROW_GAP):
        for offset, height in enumerate(CARD_ROW_HEIGHTS_PT):
            worksheet.set_row(card_top + offset, height)
    for gap_offset in range(ROW_GAP):
        worksheet.set_row(page_top + CARD_ROWS + gap_offset, ROW_GAP_HEIGHT_PT)


def set_layout(worksheet: xlsxwriter.worksheet.Worksheet) -> None:
    worksheet.set_paper(1)  # Letter
    worksheet.set_landscape()
    worksheet.hide_gridlines(2)
    worksheet.center_horizontally()
    worksheet.set_margins(left=0.25, right=0.25, top=0.25, bottom=0.25)
    worksheet.set_print_scale(100)
    set_column_widths(worksheet)


def build_formats(workbook: xlsxwriter.Workbook) -> dict[str, xlsxwriter.format.Format]:
    return {
        "header": workbook.add_format(
            {
                "bold": True,
                "font_size": 11,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "bg_color": "#D9EAD3",
            }
        ),
        "back_header": workbook.add_format(
            {
                "bold": True,
                "font_size": 11,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "bg_color": "#D9EAD3",
            }
        ),
        "label": workbook.add_format(
            {
                "bold": True,
                "font_size": 9,
                "border": 1,
                "bg_color": "#F2F2F2",
                "valign": "vcenter",
            }
        ),
        "value": workbook.add_format(
            {
                "font_size": 9,
                "border": 1,
                "valign": "vcenter",
            }
        ),
        "value_center": workbook.add_format(
            {
                "font_size": 9,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }
        ),
        "value_wrap": workbook.add_format(
            {
                "font_size": 9,
                "border": 1,
                "valign": "vcenter",
                "text_wrap": True,
            }
        ),
        "table_head": workbook.add_format(
            {
                "bold": True,
                "font_size": 9,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "bg_color": "#EDEDED",
            }
        ),
        "table_text": workbook.add_format(
            {
                "font_size": 9,
                "border": 1,
                "valign": "vcenter",
            }
        ),
        "table_center": workbook.add_format(
            {
                "font_size": 9,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }
        ),
        "status_mating": workbook.add_format(
            {
                "bold": True,
                "font_size": 9,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "bg_color": "#000000",
                "font_color": "#FFFFFF",
            }
        ),
        "status_stock": workbook.add_format(
            {
                "bold": True,
                "font_size": 9,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "bg_color": "#D9D9D9",
            }
        ),
        "note": workbook.add_format(
            {
                "font_size": 8,
                "italic": True,
                "border": 1,
                "valign": "vcenter",
            }
        ),
        "back_value": workbook.add_format(
            {
                "font_size": 8,
                "border": 1,
                "valign": "vcenter",
                "text_wrap": True,
            }
        ),
        "back_label": workbook.add_format(
            {
                "bold": True,
                "font_size": 8,
                "border": 1,
                "bg_color": "#F2F2F2",
                "valign": "vcenter",
            }
        ),
        "back_small": workbook.add_format(
            {
                "font_size": 7,
                "border": 1,
                "valign": "top",
                "text_wrap": True,
            }
        ),
    }


def write_card(
    worksheet: xlsxwriter.worksheet.Worksheet,
    start_row: int,
    start_col: int,
    cage: dict[str, Any],
    settings: dict[str, str],
    formats: dict[str, xlsxwriter.format.Format],
    include_comments: bool,
) -> None:
    disposition = safe_str(cage["disposition"]).title() or "Stock"
    status_fmt = formats["status_mating"] if disposition.lower() == "mating" else formats["status_stock"]

    visible_mice = cage["mice"][:MAX_MICE_PER_CAGE]
    overflow_count = max(0, len(cage["mice"]) - MAX_MICE_PER_CAGE)
    note_text = compact_note(cage["comment"], overflow_count=overflow_count) if include_comments else ""

    protocol_num = safe_str(cage.get("protocol_num")) or settings.get("protocol_num", "")
    approved_date = format_date_value(cage.get("approved_date")) or format_date_value(settings.get("approved_date", ""))
    expires_date = format_date_value(cage.get("expires_date")) or format_date_value(settings.get("expires_date", ""))
    source = safe_str(cage.get("source")) or settings.get("source", "")
    litter_mouseline = safe_str(cage.get("litter_mouseline"))

    worksheet.merge_range(
        start_row,
        start_col,
        start_row,
        start_col + 1,
        f"PI: {settings.get('PI_name', '')}",
        formats["header"],
    )

    worksheet.merge_range(
        start_row,
        start_col + 2,
        start_row,
        start_col + 3,
        f"Protocol: {protocol_num}",
        formats["header"],
    )

    worksheet.merge_range(
        start_row,
        start_col + 4,
        start_row,
        start_col + 5,
        f"Cage #: {cage['cage_tag']}",
        formats["header"],
    )

    worksheet.write(start_row + 1, start_col, "Contact", formats["label"])
    worksheet.merge_range(start_row + 1, start_col + 1, start_row + 1, start_col + 2, settings.get("contact_name", ""), formats["value"])
    worksheet.write(start_row + 1, start_col + 3, "Approved", formats["label"])
    worksheet.merge_range(start_row + 1, start_col + 4, start_row + 1, start_col + 5, approved_date, formats["value_center"])

    worksheet.write(start_row + 2, start_col, "Email", formats["label"])
    worksheet.merge_range(start_row + 2, start_col + 1, start_row + 2, start_col + 2, settings.get("contact_phone", ""), formats["value"])
    worksheet.write(start_row + 2, start_col + 3, "Expires", formats["label"])
    worksheet.merge_range(start_row + 2, start_col + 4, start_row + 2, start_col + 5, expires_date, formats["value_center"])

    worksheet.write(start_row + 3, start_col, "Species", formats["label"])
    worksheet.merge_range(start_row + 3, start_col + 1, start_row + 3, start_col + 2, settings.get("species", "Mouse"), formats["value"])
    worksheet.write(start_row + 3, start_col + 3, "Source", formats["label"])
    worksheet.merge_range(start_row + 3, start_col + 4, start_row + 3, start_col + 5, source, formats["value_center"])

    worksheet.write(start_row + 4, start_col, "Status", formats["label"])
    status_text = disposition.upper()
    if disposition.lower() == "mating" and safe_str(cage.get("mating_sid")):
        status_text = f"MATING #{safe_str(cage.get('mating_sid'))}"
    worksheet.merge_range(start_row + 4, start_col + 1, start_row + 4, start_col + 2, status_text, status_fmt)

    if disposition.lower() == "mating":
        worksheet.write(start_row + 4, start_col + 3, "Litter Line", formats["label"])
        worksheet.merge_range(start_row + 4, start_col + 4, start_row + 4, start_col + 5, litter_mouseline, formats["value_wrap"])
    else:
        worksheet.merge_range(start_row + 4, start_col + 3, start_row + 4, start_col + 5, "", formats["value"])

    worksheet.write(start_row + 5, start_col, "Notes", formats["label"])
    worksheet.merge_range(start_row + 5, start_col + 1, start_row + 5, start_col + 5, note_text, formats["note"])

    worksheet.write(start_row + 6, start_col + 0, "Tag", formats["table_head"])
    worksheet.write(start_row + 6, start_col + 1, "DOB", formats["table_head"])
    worksheet.write(start_row + 6, start_col + 2, "Sex", formats["table_head"])
    worksheet.merge_range(start_row + 6, start_col + 3, start_row + 6, start_col + 5, "Genotype", formats["table_head"])

    genotype_lines = list(cage["genotypes"])
    if len(genotype_lines) < len(cage["mice"]):
        genotype_lines.extend([""] * (len(cage["mice"]) - len(genotype_lines)))

    for i in range(MAX_MICE_PER_CAGE):
        row = start_row + 7 + i
        if i < len(visible_mice):
            mouse = visible_mice[i]
            genotype = normalize_dates_in_text(genotype_lines[i] if i < len(genotype_lines) else "")
            worksheet.write(row, start_col + 0, mouse["tag"], formats["table_text"])
            worksheet.write(row, start_col + 1, mouse["dob"], formats["table_center"])
            worksheet.write(row, start_col + 2, mouse["sex"], formats["table_center"])
            worksheet.merge_range(row, start_col + 3, row, start_col + 5, safe_str(genotype), formats["table_text"])
        else:
            worksheet.write_blank(row, start_col + 0, None, formats["table_text"])
            worksheet.write_blank(row, start_col + 1, None, formats["table_center"])
            worksheet.write_blank(row, start_col + 2, None, formats["table_center"])
            worksheet.merge_range(row, start_col + 3, row, start_col + 5, "", formats["table_text"])

    # Bottom spacer / border row, preserving the same physical card height.
    worksheet.merge_range(start_row + 13, start_col, start_row + 13, start_col + 5, "", formats["value"])


def _back_field(mating_info: dict[str, str] | None, key: str) -> str:
    if not mating_info:
        return ""
    return safe_str(mating_info.get(key, ""))


def write_mating_back_card(
    worksheet: xlsxwriter.worksheet.Worksheet,
    start_row: int,
    start_col: int,
    cage: dict[str, Any],
    formats: dict[str, xlsxwriter.format.Format],
) -> None:
    if safe_str(cage.get("disposition")).lower() != "mating":
        # Leave stock cage backs truly blank so duplex printing does not put any ink behind them.
        return

    mating_info = cage.get("mating_info") or {}
    cage_tag = safe_str(cage.get("cage_tag"))
    mating_sid = safe_str(cage.get("mating_sid"))

    worksheet.merge_range(
        start_row,
        start_col,
        start_row,
        start_col + 5,
        f"Mating Info Back | Cage {cage_tag} | SID {mating_sid}",
        formats["back_header"],
    )

    if not mating_info:
        worksheet.merge_range(
            start_row + 1,
            start_col,
            start_row + 5,
            start_col + 5,
            "No matching mating record was found in the uploaded mating workbook.",
            formats["back_value"],
        )
        return

    rows = [
        ("Litter Line", _back_field(mating_info, "litter_mouseline"), "Setup", _back_field(mating_info, "setup_date")),
        ("Mating Cage", _back_field(mating_info, "mating_cage_tag"), "# Litters", _back_field(mating_info, "num_litters")),
        ("Sire Tag", _back_field(mating_info, "sire_tag"), "Dam Tag", _back_field(mating_info, "dam_tag")),
        ("Sire Genotype", _back_field(mating_info, "sire_genotype"), "Dam Genotype", _back_field(mating_info, "dam_genotype")),
        ("F Mate Date", _back_field(mating_info, "female_mate_date"), "Since Setup", _back_field(mating_info, "since_setup")),
    ]

    for offset, (label1, value1, label2, value2) in enumerate(rows, start=1):
        row = start_row + offset
        worksheet.write(row, start_col, label1, formats["back_label"])
        worksheet.merge_range(row, start_col + 1, row, start_col + 2, value1, formats["back_value"])
        worksheet.write(row, start_col + 3, label2, formats["back_label"])
        worksheet.merge_range(row, start_col + 4, row, start_col + 5, value2, formats["back_value"])

    worksheet.write(start_row + 6, start_col, "Litter History", formats["back_label"])
    worksheet.merge_range(
        start_row + 6,
        start_col + 1,
        start_row + 10,
        start_col + 5,
        _back_field(mating_info, "litter_sids"),
        formats["back_small"],
    )

    worksheet.write(start_row + 11, start_col, "Comment", formats["back_label"])
    worksheet.merge_range(
        start_row + 11,
        start_col + 1,
        start_row + 13,
        start_col + 5,
        _back_field(mating_info, "comment"),
        formats["back_small"],
    )


def _rows_from_xlsx_source(xlsx_source: str | Path | bytes | BinaryIO | pd.DataFrame) -> tuple[list[list[Any]], list[str]]:
    captured_warnings: list[str] = []
    if isinstance(xlsx_source, pd.DataFrame):
        df = xlsx_source.fillna("")
        return [list(df.columns)] + df.values.tolist(), captured_warnings

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        wb = load_workbook(xlsx_source, data_only=True)

    for warning_obj in caught:
        message = str(warning_obj.message)
        if "Workbook contains no default style" not in message:
            captured_warnings.append(message)

    ws = wb.active
    return [list(r) for r in ws.iter_rows(values_only=True)], captured_warnings


def load_mating_lookup(
    mating_xlsx_source: str | Path | bytes | BinaryIO | pd.DataFrame | None,
) -> tuple[dict[str, dict[str, str]], list[str]]:
    if mating_xlsx_source is None:
        return {}, []

    rows, captured_warnings = _rows_from_xlsx_source(mating_xlsx_source)
    if not rows:
        return {}, captured_warnings

    header_index = build_header_index(rows[0], MATING_HEADER_NAMES)
    if header_index.get("mating_sid") is None:
        raise ValueError("Mating workbook is missing required column: Mating SID")

    lookup: dict[str, dict[str, str]] = {}
    for raw in rows[1:]:
        mating_sid = normalize_identifier(cell(raw, header_index, "mating_sid"))
        if not mating_sid:
            continue

        if mating_sid in lookup:
            captured_warnings.append(f"Mating SID {mating_sid} appears more than once in the mating workbook; using the last row.")

        lookup[mating_sid] = {
            "mating_sid": mating_sid,
            "mating_cage_tag": compact_multiline(cell(raw, header_index, "cage_tag"), separator=" / "),
            "comment": normalize_dates_in_text(cell(raw, header_index, "comment")),
            "litter_mouseline": safe_str(cell(raw, header_index, "litter_mouseline")),
            "setup_date": format_date_value(cell(raw, header_index, "setup_date")),
            "since_last_litter": normalize_dates_in_text(cell(raw, header_index, "since_last_litter")),
            "since_setup": normalize_dates_in_text(cell(raw, header_index, "since_setup")),
            "sire_genotype": compact_multiline(cell(raw, header_index, "sire_genotype")),
            "dam_genotype": compact_multiline(cell(raw, header_index, "dam_genotype")),
            "sire_tag": compact_multiline(cell(raw, header_index, "sire_tag"), separator=" / "),
            "female_mate_date": compact_multiline(cell(raw, header_index, "female_mate_date"), separator=" / "),
            "dam_tag": compact_multiline(cell(raw, header_index, "dam_tag"), separator=" / "),
            "num_litters": safe_str(cell(raw, header_index, "num_litters")),
            "litter_sids": compact_multiline(cell(raw, header_index, "litter_sids")),
        }

    return lookup, captured_warnings


def load_cages(xlsx_source: str | Path | bytes | BinaryIO | pd.DataFrame) -> tuple[list[dict[str, Any]], list[str]]:
    rows, captured_warnings = _rows_from_xlsx_source(xlsx_source)
    if not rows:
        return [], captured_warnings

    header_index = build_header_index(rows[0])
    required_headers = ["cage_tag", "num_mice", "mice_tags", "genotypes"]
    missing = [name for name in required_headers if header_index.get(name) is None]
    if missing:
        raise ValueError(f"Missing required column(s): {', '.join(missing)}")

    data_rows = rows[1:]
    cages: list[dict[str, Any]] = []

    for raw in data_rows:
        if header_index.get("print_card") is not None and not should_print_card(cell(raw, header_index, "print_card")):
            continue

        cage_tag = safe_str(cell(raw, header_index, "cage_tag"))
        mating_sid = normalize_identifier(cell(raw, header_index, "mating_sid"))
        litter_mouseline = safe_str(cell(raw, header_index, "litter_mouseline"))
        disposition = "mating" if mating_sid or litter_mouseline else "stock"

        mouse_lines = cleaned_lines(cell(raw, header_index, "mice_tags"))
        genotype_lines = cleaned_lines(cell(raw, header_index, "genotypes"), keep_blank_lines=True)

        if not cage_tag and not mouse_lines and not genotype_lines:
            continue

        if disposition == "mating" and not litter_mouseline:
            captured_warnings.append(
                f"Cage {cage_tag or '(blank)'} is inferred as mating, but Litter Mouseline is blank in the cage workbook."
            )

        declared_num = safe_int(cell(raw, header_index, "num_mice", 0))
        mice = parse_mouse_lines(mouse_lines)

        if declared_num and declared_num != len(mice):
            captured_warnings.append(
                f"Cage {cage_tag or '(blank)'} says {declared_num} mice, but {len(mice)} mouse-tag line(s) were found."
            )

        cages.append(
            {
                "cage_tag": cage_tag,
                "source": safe_str(cell(raw, header_index, "source")),
                "protocol_num": safe_str(cell(raw, header_index, "protocol_num")),
                "approved_date": cell(raw, header_index, "approved_date"),
                "expires_date": cell(raw, header_index, "expires_date"),
                "disposition": disposition,
                "mating_sid": mating_sid,
                "litter_mouseline": litter_mouseline,
                "mice": mice,
                "genotypes": [normalize_dates_in_text(line) for line in genotype_lines],
                "comment": normalize_dates_in_text(cell(raw, header_index, "comment")),
                "mating_info": {},
            }
        )

    return cages, captured_warnings


def attach_mating_info(
    cages: list[dict[str, Any]],
    mating_lookup: dict[str, dict[str, str]],
    warning_messages: list[str],
    mating_workbook_was_uploaded: bool,
) -> tuple[int, int]:
    matched = 0
    unmatched = 0
    for cage in cages:
        if safe_str(cage.get("disposition")).lower() != "mating":
            continue

        mating_sid = safe_str(cage.get("mating_sid"))
        mating_info = mating_lookup.get(mating_sid) if mating_sid else None
        if mating_info:
            cage["mating_info"] = mating_info
            matched += 1
            if not safe_str(cage.get("litter_mouseline")) and safe_str(mating_info.get("litter_mouseline")):
                cage["litter_mouseline"] = safe_str(mating_info.get("litter_mouseline"))
        else:
            unmatched += 1
            if mating_workbook_was_uploaded:
                warning_messages.append(
                    f"Cage {cage.get('cage_tag') or '(blank)'} uses Mating SID {mating_sid or '(blank)'}, but no matching row was found in the mating workbook."
                )
    return matched, unmatched


def write_front_page(
    worksheet: xlsxwriter.worksheet.Worksheet,
    page_top: int,
    cages: list[dict[str, Any]],
    settings: dict[str, str],
    formats: dict[str, xlsxwriter.format.Format],
    include_comments: bool,
) -> None:
    set_page_row_heights(worksheet, page_top)
    positions = [
        (page_top, 0),
        (page_top, RIGHT_CARD_START),
        (page_top + CARD_ROWS + ROW_GAP, 0),
        (page_top + CARD_ROWS + ROW_GAP, RIGHT_CARD_START),
    ]
    for cage, (row, col) in zip(cages, positions):
        write_card(worksheet, row, col, cage, settings, formats, include_comments=include_comments)


def write_back_page(
    worksheet: xlsxwriter.worksheet.Worksheet,
    page_top: int,
    cages: list[dict[str, Any]],
    formats: dict[str, xlsxwriter.format.Format],
) -> None:
    set_page_row_heights(worksheet, page_top)
    positions = [
        (page_top, 0),
        (page_top, RIGHT_CARD_START),
        (page_top + CARD_ROWS + ROW_GAP, 0),
        (page_top + CARD_ROWS + ROW_GAP, RIGHT_CARD_START),
    ]
    for cage, (row, col) in zip(cages, positions):
        write_mating_back_card(worksheet, row, col, cage, formats)


def chunked(items: list[dict[str, Any]], size: int) -> list[list[dict[str, Any]]]:
    return [items[i : i + size] for i in range(0, len(items), size)]


def build_notecards_bytes(
    xlsx_source: str | Path | bytes | BinaryIO | pd.DataFrame,
    settings: dict[str, Any] | None = None,
    include_comments: bool = True,
    mating_xlsx_source: str | Path | bytes | BinaryIO | pd.DataFrame | None = None,
) -> tuple[bytes, dict[str, Any]]:
    settings_norm = normalize_settings(settings)
    cages, warning_messages = load_cages(xlsx_source)
    mating_workbook_was_uploaded = mating_xlsx_source is not None
    mating_lookup, mating_warnings = load_mating_lookup(mating_xlsx_source)
    warning_messages.extend(mating_warnings)
    matched_mating_records, unmatched_mating_records = attach_mating_info(
        cages,
        mating_lookup,
        warning_messages,
        mating_workbook_was_uploaded=mating_workbook_was_uploaded,
    )

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    worksheet = workbook.add_worksheet("Cards")
    set_layout(worksheet)
    formats = build_formats(workbook)

    groups = chunked(cages, 4)
    include_mating_backs = mating_workbook_was_uploaded and bool(cages)

    page_top = 0
    for group in groups:
        write_front_page(worksheet, page_top, group, settings_norm, formats, include_comments=include_comments)
        page_top += ROWS_PER_SHEET
        if include_mating_backs:
            write_back_page(worksheet, page_top, group, formats)
            page_top += ROWS_PER_SHEET

    total_rows_used = page_top if cages else 0
    if total_rows_used:
        worksheet.print_area(0, 0, total_rows_used - 1, PRINT_LAST_COL)
        page_breaks = list(range(ROWS_PER_SHEET, total_rows_used, ROWS_PER_SHEET))
        if page_breaks:
            worksheet.set_h_pagebreaks(page_breaks)

    workbook.close()
    output.seek(0)

    num_front_pages = max(1, ((len(cages) - 1) // 4) + 1) if cages else 0
    metadata = {
        "num_cards": len(cages),
        "num_front_pages": num_front_pages,
        "num_pages": num_front_pages * (2 if include_mating_backs else 1),
        "warnings": warning_messages,
        "include_comments": include_comments,
        "duplex_back_cards": include_mating_backs,
        "mating_records_loaded": len(mating_lookup),
        "matched_mating_records": matched_mating_records,
        "unmatched_mating_records": unmatched_mating_records,
    }
    return output.getvalue(), metadata


def build_notecards_file(
    xlsx_source: str | Path | bytes | BinaryIO | pd.DataFrame,
    output_path: str | Path,
    settings: dict[str, Any] | None = None,
    include_comments: bool = True,
    mating_xlsx_source: str | Path | bytes | BinaryIO | pd.DataFrame | None = None,
) -> dict[str, Any]:
    content, metadata = build_notecards_bytes(
        xlsx_source=xlsx_source,
        settings=settings,
        include_comments=include_comments,
        mating_xlsx_source=mating_xlsx_source,
    )
    output_file = Path(output_path)
    output_file.write_bytes(content)
    metadata["output_path"] = str(output_file)
    return metadata


def load_settings_yaml(yaml_source: str | Path | bytes | BinaryIO) -> dict[str, str]:
    if hasattr(yaml_source, "read"):
        raw = yaml_source.read()
        if isinstance(raw, bytes):
            raw = raw.decode("utf-8")
    else:
        raw = Path(yaml_source).read_text(encoding="utf-8") if not isinstance(yaml_source, bytes) else yaml_source.decode("utf-8")
    parsed = yaml.safe_load(raw) or {}
    if not isinstance(parsed, dict):
        raise ValueError("settings.yaml must parse to a key/value mapping")
    return normalize_settings(parsed)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate mouse cage notecards from a SoftMouse cage workbook.")
    parser.add_argument("--input", default="cages.xlsx", help="Path to the cage input workbook (.xlsx)")
    parser.add_argument("--mating-input", default="", help="Optional path to the mating workbook (.xlsx) for sire/dam back cards")
    parser.add_argument("--settings-yaml", default="settings.yaml", help="Path to the YAML settings file")
    parser.add_argument("--output", default="notecards.xlsx", help="Path for the generated output workbook")
    parser.add_argument(
        "--exclude-comments",
        action="store_true",
        help="Leave the Notes row blank instead of printing spreadsheet comments",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    settings = load_settings_yaml(args.settings_yaml)
    metadata = build_notecards_file(
        xlsx_source=args.input,
        mating_xlsx_source=args.mating_input or None,
        output_path=args.output,
        settings=settings,
        include_comments=not args.exclude_comments,
    )

    print("--------------------------------------")
    print(f"Printed {metadata['num_cards']} cage card(s) in spreadsheet order.")
    print(f"Estimated physical pages: {metadata['num_pages']}")
    if metadata.get("duplex_back_cards"):
        print("Duplex back-card pages are included after each front page.")
        print(f"Mating records matched: {metadata['matched_mating_records']} / {metadata['mating_records_loaded']} loaded")
    if metadata["warnings"]:
        print("Warnings:")
        for item in metadata["warnings"]:
            print(f"- {item}")
    print(f"Saved: {metadata['output_path']}")
    print("--------------------------------------")


if __name__ == "__main__":
    main()
