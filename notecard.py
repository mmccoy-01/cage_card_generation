import argparse
import io
import re
import textwrap
import warnings
from pathlib import Path
from typing import Any, BinaryIO

import xlsxwriter
import yaml
from openpyxl import load_workbook


MAX_MICE_PER_CAGE = 6
VISIBLE_COLS_PER_CARD = 6
GUTTER_COLS = 1
CARD_ROWS = 14
ROW_GAP = 2
COL_GAP_WIDTH = 3
ROWS_PER_SHEET = CARD_ROWS * 2 + ROW_GAP
RIGHT_CARD_START = VISIBLE_COLS_PER_CARD + GUTTER_COLS
PRINT_LAST_COL = RIGHT_CARD_START + VISIBLE_COLS_PER_CARD - 1

HEADER_NAMES = {
    "cage_tag": ["cage tag"],
    "num_mice": ["# of mice", "num mice", "number of mice"],
    "disposition": ["disposition"],
    "cage_mouseline": ["cage mouseline", "mouseline", "strain"],
    "mice_tags": ["mice tags [sex, dob, age]", "mice tags", "mouse tags"],
    "genotypes": ["genotypes", "genotype"],
    "comment": ["comment", "comments", "notes"],
    "end_date": ["end date", "setup date"],
}

DEFAULT_SETTINGS = {
    "PI_name": "",
    "protocol_num": "",
    "contact_name": "",
    "contact_phone": "",
    "species": "Mouse",
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


def build_header_index(header_row: list[Any]) -> dict[str, int | None]:
    normalized = {safe_str(v).lower(): i for i, v in enumerate(header_row) if safe_str(v)}
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


def parse_mouse_lines(mouse_lines: list[str]) -> list[dict[str, str]]:
    parsed = []
    for raw in mouse_lines:
        tag = raw.split("[")[0].strip()
        sex_match = re.search(r"\[(M|F)", raw)
        dob_match = re.search(r"([0-1][0-9]-[0-3][0-9]-20[0-9]{2})", raw)
        parsed.append(
            {
                "tag": tag,
                "sex": sex_match.group(1) if sex_match else "",
                "dob": dob_match.group(1) if dob_match else "",
                "raw": raw,
            }
        )
    return parsed


def summarize_sex(mice: list[dict[str, str]]) -> str:
    males = sum(1 for m in mice if m["sex"] == "M")
    females = sum(1 for m in mice if m["sex"] == "F")
    if males and females:
        return f"{males}M / {females}F"
    if males:
        return f"{males}M"
    if females:
        return f"{females}F"
    return "-"


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


def set_layout(worksheet: xlsxwriter.worksheet.Worksheet) -> None:
    worksheet.set_paper(1)  # Letter
    worksheet.set_landscape()
    worksheet.hide_gridlines(2)
    worksheet.center_horizontally()
    worksheet.set_margins(left=0.25, right=0.25, top=0.35, bottom=0.35)
    worksheet.fit_to_pages(1, 0)

    left_widths = [8, 11, 10, 8, 9, 16]
    for i, width in enumerate(left_widths):
        worksheet.set_column(i, i, width)
    worksheet.set_column(VISIBLE_COLS_PER_CARD, VISIBLE_COLS_PER_CARD, COL_GAP_WIDTH)
    for i, width in enumerate(left_widths, start=RIGHT_CARD_START):
        worksheet.set_column(i, i, width)


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
    row_heights = [22, 18, 18, 18, 18, 18, 20, 18, 18, 18, 18, 18, 18, 18]
    for offset, height in enumerate(row_heights):
        worksheet.set_row(start_row + offset, height)

    disposition = safe_str(cage["disposition"]).title() or "Unknown"
    status_fmt = formats["status_mating"] if disposition.lower() == "mating" else formats["status_stock"]

    visible_mice = cage["mice"][:MAX_MICE_PER_CAGE]
    overflow_count = max(0, len(cage["mice"]) - MAX_MICE_PER_CAGE)
    note_text = compact_note(cage["comment"], overflow_count=overflow_count) if include_comments else ""

    worksheet.merge_range(
        start_row,
        start_col,
        start_row,
        start_col + 5,
        f"PI: {settings.get('PI_name', '')}    Protocol: {settings.get('protocol_num', '')}",
        formats["header"],
    )

    worksheet.write(start_row + 1, start_col, "Contact", formats["label"])
    worksheet.merge_range(
        start_row + 1,
        start_col + 1,
        start_row + 1,
        start_col + 5,
        settings.get("contact_name", ""),
        formats["value"],
    )

    worksheet.write(start_row + 2, start_col, "Email", formats["label"])
    worksheet.merge_range(
        start_row + 2,
        start_col + 1,
        start_row + 2,
        start_col + 5,
        settings.get("contact_phone", ""),
        formats["value"],
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
    worksheet.write(start_row + 3, start_col + 3, "Cage #", formats["label"])
    worksheet.merge_range(
        start_row + 3,
        start_col + 4,
        start_row + 3,
        start_col + 5,
        cage["cage_tag"],
        formats["value_center"],
    )

    worksheet.write(start_row + 4, start_col, "Strain", formats["label"])
    worksheet.merge_range(
        start_row + 4,
        start_col + 1,
        start_row + 4,
        start_col + 5,
        cage["mouseline"],
        formats["value_wrap"],
    )

    worksheet.write(start_row + 5, start_col, "Status", formats["label"])
    worksheet.merge_range(
        start_row + 5,
        start_col + 1,
        start_row + 5,
        start_col + 2,
        disposition.upper(),
        status_fmt,
    )
    worksheet.write(start_row + 5, start_col + 3, "Sex", formats["label"])
    worksheet.merge_range(
        start_row + 5,
        start_col + 4,
        start_row + 5,
        start_col + 5,
        summarize_sex(visible_mice),
        formats["value_center"],
    )

    worksheet.write(start_row + 6, start_col, "Notes", formats["label"])
    worksheet.merge_range(
        start_row + 6,
        start_col + 1,
        start_row + 6,
        start_col + 5,
        note_text,
        formats["note"],
    )

    worksheet.write(start_row + 7, start_col + 0, "Tag", formats["table_head"])
    worksheet.write(start_row + 7, start_col + 1, "DOB", formats["table_head"])
    worksheet.write(start_row + 7, start_col + 2, "Sex", formats["table_head"])
    worksheet.merge_range(
        start_row + 7,
        start_col + 3,
        start_row + 7,
        start_col + 5,
        "Genotype",
        formats["table_head"],
    )

    genotype_lines = list(cage["genotypes"])
    if len(genotype_lines) < len(cage["mice"]):
        genotype_lines.extend([""] * (len(cage["mice"]) - len(genotype_lines)))

    for i in range(MAX_MICE_PER_CAGE):
        row = start_row + 8 + i
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


def load_cages(xlsx_source: str | Path | bytes | BinaryIO) -> tuple[list[dict[str, Any]], list[str]]:
    captured_warnings: list[str] = []
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
    required_headers = ["cage_tag", "num_mice", "disposition", "cage_mouseline", "mice_tags", "genotypes"]
    missing = [name for name in required_headers if header_index.get(name) is None]
    if missing:
        raise ValueError(f"Missing required column(s): {', '.join(missing)}")

    data_rows = rows[1:]
    cages: list[dict[str, Any]] = []

    for raw in data_rows:
        mouseline = safe_str(cell(raw, header_index, "cage_mouseline"))
        cage_tag = safe_str(cell(raw, header_index, "cage_tag"))
        if not mouseline and not cage_tag:
            continue

        declared_num = int(cell(raw, header_index, "num_mice", 0) or 0)
        mouse_lines = cleaned_lines(cell(raw, header_index, "mice_tags"))
        mice = parse_mouse_lines(mouse_lines)
        genotype_lines = cleaned_lines(cell(raw, header_index, "genotypes"), keep_blank_lines=True)

        if declared_num and declared_num != len(mice):
            captured_warnings.append(
                f"Cage {cage_tag or '(blank)'} says {declared_num} mice, but {len(mice)} mouse-tag line(s) were found."
            )

        cages.append(
            {
                "cage_tag": cage_tag,
                "disposition": safe_str(cell(raw, header_index, "disposition")),
                "mouseline": mouseline,
                "mice": mice,
                "genotypes": genotype_lines,
                "comment": safe_str(cell(raw, header_index, "comment")),
            }
        )

    return cages, captured_warnings


def build_notecards_bytes(
    xlsx_source: str | Path | bytes | BinaryIO,
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
    xlsx_source: str | Path | bytes | BinaryIO,
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
