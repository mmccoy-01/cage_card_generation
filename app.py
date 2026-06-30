from __future__ import annotations

import pandas as pd
from shiny import App, Inputs, Outputs, Session, reactive, render, req, ui
from shiny.types import FileInfo

from notecard import build_notecards_bytes, normalize_settings, should_print_card


PRINT_CARD_COLUMN = "Print Card?"
PRINT_CARD_ALIASES = ["print card?", "print card", "print", "include", "include card", "selected"]

EDITABLE_EXTRA_COLUMNS = {
    "Protocol": ["protocol", "protocol number", "protocol #", "protocol num"],
    "Approved": ["approved", "approved date", "approval date"],
    "Expires": ["expires", "expires date", "expiration date", "expiry date"],
    "Source": ["source"],
}


def _normalized_column_name(value: object) -> str:
    return " ".join(str(value).strip().lower().replace("_", " ").split())


def _find_equivalent_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
    normalized_aliases = {_normalized_column_name(alias) for alias in aliases}
    for col in df.columns:
        if _normalized_column_name(col) in normalized_aliases:
            return str(col)
    return None


def ensure_editable_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add editable card-specific columns if the uploaded workbook does not
    already contain equivalent columns. Print Card? is always kept as the
    far-left boolean column so Shiny renders it as checkbox-style values.
    """
    df = df.copy().fillna("")

    existing_print_col = _find_equivalent_column(df, PRINT_CARD_ALIASES)
    if existing_print_col is None:
        df.insert(0, PRINT_CARD_COLUMN, True)
    else:
        df[existing_print_col] = df[existing_print_col].map(should_print_card).astype(bool)
        if existing_print_col != PRINT_CARD_COLUMN:
            df = df.rename(columns={existing_print_col: PRINT_CARD_COLUMN})
        # Move Print Card? to the far left.
        cols = [PRINT_CARD_COLUMN] + [col for col in df.columns if col != PRINT_CARD_COLUMN]
        df = df[cols]

    existing = {_normalized_column_name(col) for col in df.columns}
    for canonical_name, aliases in EDITABLE_EXTRA_COLUMNS.items():
        if not any(_normalized_column_name(alias) in existing for alias in aliases):
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
        ui.download_button("download_cards", "Download notecards.xlsx", class_="btn-primary"),
        width=380,
    ),
    ui.h2("Mouse cage card generator"),
    ui.p(
        "Upload a cage workbook, optionally upload a mating workbook, edit the table if needed, "
        "uncheck cages you do not want to print, and download a print-ready notecards.xlsx file."
    ),
    ui.h4("Status"),
    ui.output_text_verbatim("status"),
    ui.h4("Editable uploaded cage sheet"),
    ui.p(
        "Use the Print Card? checkbox column to select/deselect cages. "
        "Blank per-cage Protocol, Approved, Expires, or Source values fall back to the sidebar defaults."
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
            edit_version.set(0)
            return

        try:
            df = pd.read_excel(file_info["datapath"])
            editable_df.set(ensure_editable_columns(df))
            # Do not read edit_version() inside this reactive effect. Reading and
            # then setting the same reactive value creates a self-invalidating
            # loop, which makes the app buffer until the server disconnects.
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

        col_name = df.columns[patch["column_index"]]
        value = "" if patch["value"] is None else patch["value"]
        if col_name == PRINT_CARD_COLUMN:
            value = should_print_card(value)

        # Mutate the backing dataframe in place and only invalidate downstream
        # generation. Avoid editable_df.set(df) here, because that re-renders the
        # whole DataGrid and can jump the user back to the top while editing.
        df.iat[patch["row_index"], patch["column_index"]] = value
        with reactive.isolate():
            edit_version.set(edit_version() + 1)

        return value

    @reactive.calc
    def generation_result() -> dict[str, object] | None:
        file_info = uploaded_file()
        if file_info is None:
            return None

        df = editable_df()
        edit_version()
        if df.empty:
            return None

        if "error" in df.columns:
            return {"content": None, "metadata": None, "error": str(df["error"].iloc[0])}

        mating_file_info = uploaded_mating_file()
        mating_source = None if mating_file_info is None else mating_file_info["datapath"]

        try:
            content, metadata = build_notecards_bytes(
                xlsx_source=df,
                mating_xlsx_source=mating_source,
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
            return "Upload a cage .xlsx file to begin. Upload a mating .xlsx file only if you want sire/dam backs."

        result = generation_result()
        if result is None:
            return "Upload a cage .xlsx file to begin."

        if result["error"]:
            return f"Could not build cards: {result['error']}"

        metadata = result["metadata"]
        assert isinstance(metadata, dict)

        lines = [
            f"Ready: {metadata['num_cards']} selected card(s), {metadata['num_pages']} physical print page(s).",
            f"Front pages: {metadata['num_front_pages']}",
            f"Comments included: {'yes' if metadata['include_comments'] else 'no'}",
            "Using edited table values.",
        ]

        if metadata.get("duplex_back_cards"):
            lines.extend(
                [
                    "Mating backs: enabled. Front pages alternate with matching back pages for duplex printing.",
                    "Stock cage backs are left blank so the next cage front will not print behind them.",
                    f"Mating records loaded: {metadata['mating_records_loaded']}",
                    f"Mating records matched to selected cages: {metadata['matched_mating_records']}",
                    "Printer note: choose double-sided printing in the print dialog; Excel files cannot reliably force that printer setting on every computer.",
                ]
            )
        else:
            lines.append("Mating backs: disabled because no mating workbook is uploaded.")

        warnings_list = metadata.get("warnings", [])
        if warnings_list:
            lines.append("Warnings:")
            lines.extend(f"- {item}" for item in warnings_list)

        return "\n".join(lines)

    @render.download(filename="notecards.xlsx")
    def download_cards():
        result = generation_result()
        req(result is not None)
        req(result["error"] is None)

        content = result["content"]
        req(isinstance(content, (bytes, bytearray)))

        yield content


app = App(app_ui, server)
