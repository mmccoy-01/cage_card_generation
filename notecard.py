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
from openpyxl.utils.datetime import from_excel


MAX_MICE_PER_CAGE = 6
VISIBLE_COLS_PER_CARD = 6
GUTTER_COLS = 1

# Physical cage-card target: 12.7 cm length x 7.7 cm height
# 12.7 cm is exactly 5 inches. Excel uses row heights in points and
# xlsxwriter can set column widths in pixels. At 96 px/in, 5 in = 480 px.
CARD_WIDTH_CM = 12.7
CARD_HEIGHT_CM = 7.7
CARD_WIDTH_PIXELS = round((CARD_WIDTH_CM / 2.54) * 96)
CARD_HEIGHT_POINTS = (CARD_HEIGHT_CM / 2.54) * 72
GUTTER_WIDTH_PIXELS = round(0.25 * 96)

# The visible card uses 13 rows: 7 header/detail rows + 6 mouse rows.
CARD_ROWS = 13
ROW_GAP = 2
ROWS_PER_SHEET = CARD_ROWS * 2 + ROW_GAP
RIGHT_CARD_START = VISIBLE_COLS_PER_CARD + GUTTER_COLS
PRINT_LAST_COL = RIGHT_CARD_START + VISIBLE_COLS_PER_CARD - 1

HEADER_NAMES = {
    "cage_tag": ["cage tag"],
    "num_mice": ["# of mice", "num mice", "number of mice"],
    # SoftMouse no longer exports Disposition or Cage Mouseline for this workflow.
    # Card status is inferred from mating/litter fields below.
    "disposition": ["disposition"],  # optional backward-compatibility alias only
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


def safe_str(value: Any) -> str:
    return "" if value is None else str(value).strip()


def normalize_settings(settings: dict[str, Any] | None) -> dict[str, str]:
    normalized = dict(DEFAULT_SETTINGS)
    if settings:
        normalized.update({k: safe_str(v) for k, v in settings.items()})
    if not normalized["species"]:
        normalized["species"] = "Mouse"
    return normalized


def normalize_header(value: Any) -> str:
    text = safe_str(value).lower()
    text = text.replace("\n", " ").replace("\r", " ").replace("_", " ")
    text = re.sub(r"^\$+\s*", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def build_header_index(header_row: list[Any]) -> dict[str, int | None]:
    normalized = {normalize_header(v): i for i, v in enumerate(header_row) if safe_str(v)}
    out: dict[str, int | None] = {}
    for key, candidates in HEADER_NAMES.items():
        out[key] = None
        for name in candidates:
            if name in normalized:
                out[key] = normalized[name]
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


def format_date_value(value: Any) -> str:
    """Return date-like values as YYYY-MM-DD while leaving non-dates unchanged."""
    if value is None:
        return ""

    if isinstance(value, pd.Timestamp):
        if pd.isna(value):
            return ""
        return value.strftime("%Y-%m-%d")

    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if pd.isna(value):
            return ""
        # Excel serial dates are usually in this range for modern colony records.
        if 20000 <= float(value) <= 60000:
            try:
                return from_excel(value).strftime("%Y-%m-%d")
            except Exception:
                pass
        return safe_str(value)

    text = safe_str(value)
    if not text:
        return ""

    parsed = pd.to_datetime(text, errors="coerce")
    if pd.notna(parsed):
        return parsed.strftime("%Y-%m-%d")

    return text


def extract_date_text(value: str) -> str:
    """Extract a date token from mouse-tag text and normalize it to YYYY-MM-DD."""
    patterns = [
        r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b",
        r"\b\d{1,2}[-/]\d{1,2}[-/](?:\d{2}|\d{4})\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, value)
        if match:
            return format_date_value(match.group(0))
    return ""


def parse_mouse_lines(mouse_lines: list[str]) -> list[dict[str, str]]:
    parsed = []
    for raw in mouse_lines:
        tag = raw.split("[")[0].strip()
        sex_match = re.search(r"\[(M|F)", raw, flags=re.IGNORECASE)
        parsed.append(
            {
                "tag": tag,
                "sex": sex_match.group(1).upper() if sex_match else "",
                "dob": extract_date_text(raw),
                "raw": raw,
            }
        )
    return parsed


def compact_note(comment: str, overflow_count: int = 0) -> str:
    parts = []
    note = " ".join(cleaned_lines(comment))
    if note:
        parts.append(note)
    if overflow_count > 0:
        parts.append(f"+{overflow_count} more mouse(s) not shown")
    if not parts:
        return ""
    return textwrap.shorten(" | ".join(parts), width=78, placeholder="...")


def card_column_widths_pixels() -> list[int]:
    """Return six column widths whose total is the 12.7 cm card width."""
    proportions = [8, 11, 10, 8, 9, 16]
    raw = [CARD_WIDTH_PIXELS * value / sum(proportions) for value in proportions]
    widths = [round(value) for value in raw]
    widths[-1] += CARD_WIDTH_PIXELS - sum(widths)
    return widths


def card_row_heights_points() -> list[float]:
    """Return thirteen row heights whose total is the 7.7 cm card height."""
    proportions = [22, 18, 18, 18, 18, 20, 18, 18, 18, 18, 18, 18, 18]
    scale = CARD_HEIGHT_POINTS / sum(proportions)
    heights = [value * scale for value in proportions]
    heights[-1] += CARD_HEIGHT_POINTS - sum(heights)
    return heights




def set_column_width_pixels(worksheet: xlsxwriter.worksheet.Worksheet, first_col: int, last_col: int, width: int) -> None:
    """Set a column width in pixels, with a fallback for older xlsxwriter."""
    if hasattr(worksheet, "set_column_pixels"):
        worksheet.set_column_pixels(first_col, last_col, width)
    else:
        # Older xlsxwriter only supports Excel character-width units. This is
        # approximate, but keeps the file usable if set_column_pixels is absent.
        worksheet.set_column(first_col, last_col, width / 7)


def set_layout(worksheet: xlsxwriter.worksheet.Worksheet) -> None:
    worksheet.set_paper(1)  # Letter
    worksheet.set_landscape()
    worksheet.hide_gridlines(2)
    worksheet.center_horizontally()
    worksheet.set_margins(left=0.25, right=0.25, top=0.35, bottom=0.35)

    # Keep print scale at 100% so the physical card dimensions are preserved.
    # Two 12.7 cm cards plus a 0.25 in gutter fit on letter landscape paper.
    worksheet.set_print_scale(100)

    left_widths = card_column_widths_pixels()
    for i, width in enumerate(left_widths):
        set_column_width_pixels(worksheet, i, i, width)
    set_column_width_pixels(worksheet, VISIBLE_COLS_PER_CARD, VISIBLE_COLS_PER_CARD, GUTTER_WIDTH_PIXELS)
    for i, width in enumerate(left_widths, start=RIGHT_CARD_START):
        set_column_width_pixels(worksheet, i, i, width)


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
    row_heights = card_row_heights_points()
    for offset, height in enumerate(row_heights):
        worksheet.set_row(start_row + offset, height)

    disposition = safe_str(cage["disposition"]).title() or "Unknown"
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
    worksheet.merge_range(
        start_row + 1,
        start_col + 1,
        start_row + 1,
        start_col + 2,
        settings.get("contact_name", ""),
        formats["value"],
    )

    worksheet.write(start_row + 1, start_col + 3, "Approved", formats["label"])
    worksheet.merge_range(
        start_row + 1,
        start_col + 4,
        start_row + 1,
        start_col + 5,
        approved_date,
        formats["value_center"],
    )

    worksheet.write(start_row + 2, start_col, "Email", formats["label"])
    worksheet.merge_range(
        start_row + 2,
        start_col + 1,
        start_row + 2,
        start_col + 2,
        settings.get("contact_phone", ""),
        formats["value"],
    )

    worksheet.write(start_row + 2, start_col + 3, "Expires", formats["label"])
    worksheet.merge_range(
        start_row + 2,
        start_col + 4,
        start_row + 2,
        start_col + 5,
        expires_date,
        formats["value_center"],
    )

    worksheet.write(start_row + 3, start_col, "Species", formats["label"])
    worksheet.merge_range(
        start_row + 3,
        start_col + 1,
        start_row + 3,
        start_col + 2,
        settings.get("species", "Mouse"),
        formats["value"],
    )

    worksheet.write(start_row + 3, start_col + 3, "Source", formats["label"])
    worksheet.merge_range(
        start_row + 3,
        start_col + 4,
        start_row + 3,
        start_col + 5,
        source,
        formats["value_center"],
    )

    worksheet.write(start_row + 4, start_col, "Status", formats["label"])
    worksheet.merge_range(
        start_row + 4,
        start_col + 1,
        start_row + 4,
        start_col + 2,
        disposition.upper(),
        status_fmt,
    )

    if disposition.lower() == "mating":
        litter_text = f"Litter Mouseline: {litter_mouseline}" if litter_mouseline else "Litter Mouseline:"
        worksheet.merge_range(
            start_row + 4,
            start_col + 3,
            start_row + 4,
            start_col + 5,
            litter_text,
            formats["value_wrap"],
        )
    else:
        worksheet.merge_range(
            start_row + 4,
            start_col + 3,
            start_row + 4,
            start_col + 5,
            "",
            formats["value"],
        )

    worksheet.write(start_row + 5, start_col, "Notes", formats["label"])
    worksheet.merge_range(
        start_row + 5,
        start_col + 1,
        start_row + 5,
        start_col + 5,
        note_text,
        formats["note"],
    )

    worksheet.write(start_row + 6, start_col + 0, "Tag", formats["table_head"])
    worksheet.write(start_row + 6, start_col + 1, "DOB", formats["table_head"])
    worksheet.write(start_row + 6, start_col + 2, "Sex", formats["table_head"])
    worksheet.merge_range(
        start_row + 6,
        start_col + 3,
        start_row + 6,
        start_col + 5,
        "Genotype",
        formats["table_head"],
    )

    genotype_lines = list(cage["genotypes"])
    if len(genotype_lines) < len(cage["mice"]):
        genotype_lines.extend([""] * (len(cage["mice"]) - len(genotype_lines)))

    for i in range(MAX_MICE_PER_CAGE):
        row = start_row + 7 + i
        if i < len(visible_mice):
            mouse = visible_mice[i]
            genotype = genotype_lines[i] if i < len(genotype_lines) else ""
            worksheet.write(row, start_col + 0, mouse["tag"], formats["table_text"])
            worksheet.write(row, start_col + 1, mouse["dob"], formats["table_center"])
            worksheet.write(row, start_col + 2, mouse["sex"], formats["table_center"])
            worksheet.merge_range(
                row,
                start_col + 3,
                row,
                start_col + 5,
                safe_str(genotype),
                formats["table_text"],
            )
        else:
            worksheet.write_blank(row, start_col + 0, None, formats["table_text"])
            worksheet.write_blank(row, start_col + 1, None, formats["table_center"])
            worksheet.write_blank(row, start_col + 2, None, formats["table_center"])
            worksheet.merge_range(
                row,
                start_col + 3,
                row,
                start_col + 5,
                "",
                formats["table_text"],
            )


def load_cages(xlsx_source: str | Path | bytes | BinaryIO | pd.DataFrame) -> tuple[list[dict[str, Any]], list[str]]:
    captured_warnings: list[str] = []

    if isinstance(xlsx_source, pd.DataFrame):
        df = xlsx_source.fillna("")
        rows = [list(df.columns)] + df.values.tolist()
    else:
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            wb = load_workbook(xlsx_source, data_only=True)

        for warning_obj in caught:
            message = str(warning_obj.message)
            if "Workbook contains no default style" not in message:
                captured_warnings.append(message)

        ws = wb.active
        rows = [list(r) for r in ws.iter_rows(values_only=True)]

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
        cage_tag = safe_str(cell(raw, header_index, "cage_tag"))
        mating_sid = safe_str(cell(raw, header_index, "mating_sid"))
        litter_mouseline = safe_str(cell(raw, header_index, "litter_mouseline"))
        disposition = "mating" if mating_sid or litter_mouseline else "stock"

        declared_num = safe_int(cell(raw, header_index, "num_mice", 0))
        mouse_lines = cleaned_lines(cell(raw, header_index, "mice_tags"))
        mice = parse_mouse_lines(mouse_lines)
        genotype_lines = cleaned_lines(cell(raw, header_index, "genotypes"), keep_blank_lines=True)

        if not any([cage_tag, mating_sid, litter_mouseline, mouse_lines, genotype_lines]):
            continue

        if disposition == "mating" and not litter_mouseline:
            captured_warnings.append(
                f"Cage {cage_tag or '(blank)'} has mating information, but Litter Mouseline is blank or missing."
            )

        if declared_num and declared_num != len(mice):
            captured_warnings.append(
                f"Cage {cage_tag or '(blank)'} says {declared_num} mice, but {len(mice)} mouse-tag line(s) were found."
            )

        cages.append(
            {
                "cage_tag": cage_tag,
                "source": safe_str(cell(raw, header_index, "source")),
                "protocol_num": safe_str(cell(raw, header_index, "protocol_num")),
                "approved_date": safe_str(cell(raw, header_index, "approved_date")),
                "expires_date": safe_str(cell(raw, header_index, "expires_date")),
                "disposition": disposition,
                "mating_sid": mating_sid,
                "litter_mouseline": litter_mouseline,
                "mice": mice,
                "genotypes": genotype_lines,
                "comment": safe_str(cell(raw, header_index, "comment")),
            }
        )

    return cages, captured_warnings


def build_notecards_bytes(
    xlsx_source: str | Path | bytes | BinaryIO | pd.DataFrame,
    settings: dict[str, Any] | None = None,
    include_comments: bool = True,
) -> tuple[bytes, dict[str, Any]]:
    settings_norm = normalize_settings(settings)
    cages, warning_messages = load_cages(xlsx_source)

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    worksheet = workbook.add_worksheet("Cards")
    set_layout(worksheet)
    formats = build_formats(workbook)

    current_sheet_top = 0
    slot_on_sheet = 0
    page_breaks: list[int] = []

    for cage in cages:
        card_row_offset = 0 if slot_on_sheet < 2 else CARD_ROWS + ROW_GAP
        card_col = 0 if slot_on_sheet % 2 == 0 else RIGHT_CARD_START
        write_card(
            worksheet,
            current_sheet_top + card_row_offset,
            card_col,
            cage,
            settings_norm,
            formats,
            include_comments=include_comments,
        )

        slot_on_sheet += 1
        if slot_on_sheet == 4:
            current_sheet_top += ROWS_PER_SHEET
            page_breaks.append(current_sheet_top)
            slot_on_sheet = 0

    total_rows_used = current_sheet_top + (
        ROWS_PER_SHEET if slot_on_sheet == 0 and cages else CARD_ROWS + (CARD_ROWS + ROW_GAP if slot_on_sheet > 2 else 0)
    )
    worksheet.print_area(0, 0, max(ROWS_PER_SHEET - 1, total_rows_used - 1), PRINT_LAST_COL)
    if page_breaks:
        worksheet.set_h_pagebreaks(sorted(set(page_breaks)))

    workbook.close()
    output.seek(0)

    metadata = {
        "num_cards": len(cages),
        "num_pages": max(1, ((len(cages) - 1) // 4) + 1) if cages else 0,
        "warnings": warning_messages,
        "include_comments": include_comments,
    }
    return output.getvalue(), metadata


def build_notecards_file(
    xlsx_source: str | Path | bytes | BinaryIO | pd.DataFrame,
    output_path: str | Path,
    settings: dict[str, Any] | None = None,
    include_comments: bool = True,
) -> dict[str, Any]:
    content, metadata = build_notecards_bytes(
        xlsx_source=xlsx_source,
        settings=settings,
        include_comments=include_comments,
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
    parser = argparse.ArgumentParser(description="Generate mouse cage notecards from a SoftMouse workbook.")
    parser.add_argument("--input", default="softmousedb.xlsx", help="Path to the input SoftMouse workbook (.xlsx)")
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
        output_path=args.output,
        settings=settings,
        include_comments=not args.exclude_comments,
    )

    print("--------------------------------------")
    print(f"Printed {metadata['num_cards']} cage card(s) in spreadsheet order.")
    print(f"Estimated pages: {metadata['num_pages']}")
    if metadata["warnings"]:
        print("Warnings:")
        for item in metadata["warnings"]:
            print(f"- {item}")
    print(f"Saved: {metadata['output_path']}")
    print("--------------------------------------")


if __name__ == "__main__":
    main()
