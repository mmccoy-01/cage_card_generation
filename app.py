from __future__ import annotations

import pandas as pd
from shiny import App, Inputs, Outputs, Session, reactive, render, req, ui
from shiny.types import FileInfo

from notecard import PRINT_COLUMN, build_notecards_pdf_bytes, normalize_settings, safe_bool


EDITABLE_EXTRA_COLUMNS = {
    "Protocol": ["protocol", "protocol number", "protocol #", "protocol num"],
    "Approved": ["approved", "approved date", "approval date"],
    "Expires": ["expires", "expires date", "expiration date", "expiry date"],
    "Source": ["source"],
}


def _normalized_column_name(value: object) -> str:
    return " ".join(str(value).strip().lower().split())


def ensure_editable_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add app-specific editable columns if the uploaded workbook does not already
    contain equivalent columns. Print Card? is intentionally first and defaults
    to True so users uncheck cages they do not want to print.
    """
    df = df.copy().fillna("")
    normalized_to_actual = {_normalized_column_name(col): col for col in df.columns}

    existing_print_col = None
    for candidate in ["print card?", "print card", "print", "include", "selected"]:
        if candidate in normalized_to_actual:
            existing_print_col = normalized_to_actual[candidate]
            break

    if existing_print_col is None:
        df.insert(0, PRINT_COLUMN, True)
    else:
        df[existing_print_col] = df[existing_print_col].map(lambda value: safe_bool(value, default=True)).astype(bool)
        if existing_print_col != PRINT_COLUMN:
            df = df.rename(columns={existing_print_col: PRINT_COLUMN})
        cols = [PRINT_COLUMN] + [col for col in df.columns if col != PRINT_COLUMN]
        df = df[cols]

    existing = {_normalized_column_name(col) for col in df.columns}
    for canonical_name, aliases in EDITABLE_EXTRA_COLUMNS.items():
        if not any(alias in existing for alias in aliases):
            df[canonical_name] = ""

    return df


app_ui = ui.page_sidebar(
    ui.sidebar(
        ui.h4("Inputs"),
        ui.input_file(
            "softmouse_file",
            "Cage workbook (.xlsx)",
            accept=[".xlsx"],
            multiple=False,
        ),
        ui.input_file(
            "mating_file",
            "Mating workbook for sire/dam backs (.xlsx, optional)",
            accept=[".xlsx"],
            multiple=False,
        ),
        ui.hr(),
        ui.input_text("pi_name", "PI name", ""),
        ui.input_text("protocol_num", "Default protocol number", ""),
        ui.input_text("approved_date", "Default approved date", ""),
        ui.input_text("expires_date", "Default expires date", ""),
        ui.input_text("contact_name", "Contact name", ""),
        ui.input_text("contact_phone", "Contact email / phone", ""),
        ui.input_text("species", "Species", "Mouse"),
        ui.input_text("source", "Source", ""),
        ui.input_checkbox("include_comments", "Include comments on cards", value=True),
        ui.hr(),
        ui.download_button("download_cards", "Download notecards.pdf", class_="btn-primary"),
        width=360,
    ),
    ui.h2("Mouse cage card generator"),
    ui.p(
        "Upload a cage workbook, optionally upload a mating workbook, edit the table if needed, "
        "uncheck cages you do not want to print, and download a print-ready PDF."
    ),
    ui.h4("Status"),
    ui.output_text_verbatim("status"),
    ui.h4("Editable uploaded cage sheet"),
    ui.p(
        "Use the Print Card? checkbox column to select/deselect cages. Blank per-cage Protocol, "
        "Approved, Expires, or Source values fall back to the sidebar defaults."
    ),
    ui.output_data_frame("editable_sheet"),
    title="Mouse cage cards",
)


def server(input: Inputs, output: Outputs, session: Session):
    editable_df: reactive.Value[pd.DataFrame] = reactive.Value(pd.DataFrame())
    edit_version: reactive.Value[int] = reactive.Value(0)

    @reactive.calc
    def uploaded_file() -> FileInfo | None:
        files: list[FileInfo] | None = input.softmouse_file()
        return None if not files else files[0]

    @reactive.calc
    def uploaded_mating_file() -> FileInfo | None:
        files: list[FileInfo] | None = input.mating_file()
        return None if not files else files[0]

    @reactive.calc
    def settings() -> dict[str, str]:
        return normalize_settings(
            {
                "PI_name": input.pi_name(),
                "protocol_num": input.protocol_num(),
                "approved_date": input.approved_date(),
                "expires_date": input.expires_date(),
                "contact_name": input.contact_name(),
                "contact_phone": input.contact_phone(),
                "species": input.species(),
                "source": input.source(),
            }
        )

    @reactive.effect
    def _load_uploaded_workbook() -> None:
        file_info = uploaded_file()

        if file_info is None:
            editable_df.set(pd.DataFrame())
            return

        try:
            df = pd.read_excel(file_info["datapath"])
            editable_df.set(ensure_editable_columns(df))
            edit_version.set(0)
        except Exception as exc:
            editable_df.set(pd.DataFrame({"error": [str(exc)]}))
            edit_version.set(0)

    @render.data_frame
    def editable_sheet():
        df = editable_df()

        if df.empty:
            return render.DataGrid(
                pd.DataFrame(),
                editable=False,
                filters=False,
                height="500px",
                width="100%",
            )

        return render.DataGrid(
            df,
            editable=True if "error" not in df.columns else False,
            filters=False,
            height="500px",
            width="100%",
        )

    @editable_sheet.set_patch_fn
    def _update_editable_sheet(patch):
        df = editable_df()

        if df.empty or "error" in df.columns:
            return patch["value"]

        value = "" if patch["value"] is None else patch["value"]
        col_name = df.columns[patch["column_index"]]
        if col_name == PRINT_COLUMN:
            value = safe_bool(value, default=True)

        # Mutate in place and invalidate downstream generation without forcing
        # the whole DataGrid to rebuild after every edit.
        df.iat[patch["row_index"], patch["column_index"]] = value
        edit_version.set(edit_version() + 1)

        return value

    @reactive.calc
    def generation_result() -> dict[str, object] | None:
        file_info = uploaded_file()
        if file_info is None:
            return None

        _ = edit_version()
        df = editable_df()
        if df.empty:
            return None

        if "error" in df.columns:
            return {"content": None, "metadata": None, "error": str(df["error"].iloc[0])}

        try:
            mating_info = uploaded_mating_file()
            mating_source = mating_info["datapath"] if mating_info is not None else None
            content, metadata = build_notecards_pdf_bytes(
                xlsx_source=df,
                mating_source=mating_source,
                settings=settings(),
                include_comments=input.include_comments(),
            )
            return {"content": content, "metadata": metadata, "error": None}
        except Exception as exc:
            return {"content": None, "metadata": None, "error": str(exc)}

    @render.text
    def status() -> str:
        file_info = uploaded_file()
        if file_info is None:
            return "Upload a cage .xlsx file to begin."

        result = generation_result()
        if result is None:
            return "Upload a cage .xlsx file to begin."

        if result["error"]:
            return f"Could not build cards: {result['error']}"

        metadata = result["metadata"]
        assert isinstance(metadata, dict)

        lines = [
            f"Ready: {metadata['num_cards']} card(s) in a {metadata['num_pages']}-page PDF.",
            f"Card size: {metadata['card_width_mm']} mm wide x {metadata['card_height_mm']} mm tall.",
            f"Comments included: {'yes' if metadata['include_comments'] else 'no'}",
            "PDF layout is duplex-safe: each front page is followed by its matching back page; stock cage backs are blank.",
            "Print at Actual Size / 100% scale, double-sided, flip on long edge.",
        ]

        warnings_list = metadata.get("warnings", [])
        if warnings_list:
            lines.append("Warnings:")
            lines.extend(f"- {item}" for item in warnings_list)

        return "\n".join(lines)

    @render.download(filename="notecards.pdf")
    def download_cards():
        result = generation_result()
        req(result is not None)
        req(result["error"] is None)

        content = result["content"]
        req(isinstance(content, (bytes, bytearray)))

        yield content


app = App(app_ui, server)
