from __future__ import annotations

import pandas as pd
from shiny import App, Inputs, Outputs, Session, reactive, render, req, ui
from shiny.types import FileInfo

from notecard import build_notecards_bytes, normalize_settings


PRINT_COLUMN = "Print Card?"
PRINT_COLUMN_ALIASES = [
    "print card?",
    "print card",
    "print?",
    "print",
    "include card?",
    "include card",
    "include?",
    "include",
]

EDITABLE_EXTRA_COLUMNS = {
    "Protocol": ["protocol", "protocol number", "protocol #", "protocol num"],
    "Approved": ["approved", "approved date", "approval date"],
    "Expires": ["expires", "expires date", "expiration date", "expiry date"],
    "Source": ["source"],
}


def _normalized_column_name(value: object) -> str:
    text = str(value).strip().lower().replace("_", " ")
    text = text.lstrip("$").strip()
    return " ".join(text.split())


def _coerce_print_value(value: object, default: bool = True) -> bool:
    if value is None:
        return default

    try:
        if pd.isna(value):
            return default
    except TypeError:
        pass

    if isinstance(value, bool):
        return value

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return bool(value)

    text = str(value).strip().lower()
    if not text:
        return default

    if text in {"true", "t", "yes", "y", "1", "x", "checked", "print", "include", "selected"}:
        return True
    if text in {"false", "f", "no", "n", "0", "unchecked", "skip", "exclude", "omit"}:
        return False

    return default


def ensure_editable_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add app-specific editable columns if the uploaded workbook does not
    already contain an equivalent column.
    """
    df = df.copy()

    normalized_to_original = {_normalized_column_name(col): col for col in df.columns}

    print_col = None
    for alias in PRINT_COLUMN_ALIASES:
        if alias in normalized_to_original:
            print_col = normalized_to_original[alias]
            break

    if print_col is None:
        df.insert(0, PRINT_COLUMN, True)
    else:
        if print_col != PRINT_COLUMN:
            df = df.rename(columns={print_col: PRINT_COLUMN})

        # Keep the checkbox/toggle column at the far left even if the uploaded
        # workbook already had an equivalent print/include column elsewhere.
        remaining_cols = [col for col in df.columns if col != PRINT_COLUMN]
        df = df[[PRINT_COLUMN, *remaining_cols]]

    existing = {_normalized_column_name(col) for col in df.columns}
    for canonical_name, aliases in EDITABLE_EXTRA_COLUMNS.items():
        if not any(alias in existing for alias in aliases):
            df[canonical_name] = ""

    df = df.fillna("")

    # Store this as real boolean dtype so Shiny renders it as a checkbox-style
    # editable column instead of plain TRUE/FALSE text.
    df[PRINT_COLUMN] = (
        df[PRINT_COLUMN]
        .apply(lambda value: _coerce_print_value(value, default=True))
        .astype(bool)
    )
    return df


def selected_rows_for_print(df: pd.DataFrame) -> pd.DataFrame:
    """Return only rows selected for printing; remove the UI-only toggle column."""
    if df.empty or PRINT_COLUMN not in df.columns:
        return df.copy()

    mask = df[PRINT_COLUMN].apply(lambda value: _coerce_print_value(value, default=True))
    return df.loc[mask].drop(columns=[PRINT_COLUMN], errors="ignore").copy()


app_ui = ui.page_sidebar(
    ui.sidebar(
        ui.h4("Inputs"),
        ui.input_file(
            "softmouse_file",
            "SoftMouse workbook (.xlsx)",
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
        width=360,
    ),
    ui.h2("Mouse cage card generator"),
    ui.p(
        "Upload a SoftMouse workbook, choose which cages to print, edit the table if needed, "
        "and download a print-ready notecards.xlsx file."
    ),
    ui.h4("Status"),
    ui.output_text_verbatim("status"),
    ui.h4("Editable uploaded sheet"),
    ui.p(
        "Use the leftmost Print Card? checkbox column to select/deselect cages; all rows "
        "start selected by default. Blank per-cage Protocol, Approved, Expires, or Source "
        "values fall back to the sidebar defaults."
    ),
    ui.output_data_frame("editable_sheet"),
    title="Mouse cage cards",
)


def server(input: Inputs, output: Outputs, session: Session):
    editable_df: reactive.Value[pd.DataFrame] = reactive.Value(pd.DataFrame())
    edit_version: reactive.Value[int] = reactive.Value(0)
    edit_version_counter = {"value": 0}

    def bump_edit_version() -> None:
        edit_version_counter["value"] += 1
        edit_version.set(edit_version_counter["value"])

    @reactive.calc
    def uploaded_file() -> FileInfo | None:
        files: list[FileInfo] | None = input.softmouse_file()
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
            bump_edit_version()
            return

        try:
            df = pd.read_excel(file_info["datapath"])
            editable_df.set(ensure_editable_columns(df))
            bump_edit_version()
        except Exception as exc:
            editable_df.set(pd.DataFrame({"error": [str(exc)]}))
            bump_edit_version()

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
        if col_name == PRINT_COLUMN:
            value = _coerce_print_value(patch["value"], default=False)
        else:
            value = "" if patch["value"] is None else patch["value"]

        # Mutate in place instead of resetting the whole dataframe. This keeps the
        # displayed grid from jumping back to the top after each edit.
        df.iat[patch["row_index"], patch["column_index"]] = value
        bump_edit_version()

        return value

    @reactive.calc
    def generation_result() -> dict[str, object] | None:
        file_info = uploaded_file()
        if file_info is None:
            return None

        df = editable_df()
        edit_version()  # Recompute after cell edits without forcing a grid redraw.
        if df.empty:
            return None

        if "error" in df.columns:
            return {"content": None, "metadata": None, "error": str(df["error"].iloc[0])}

        try:
            selected_df = selected_rows_for_print(df)
            content, metadata = build_notecards_bytes(
                xlsx_source=selected_df,
                settings=settings(),
                include_comments=input.include_comments(),
            )
            metadata["num_uploaded_rows"] = int(len(df))
            metadata["num_selected_rows"] = int(len(selected_df))
            return {"content": content, "metadata": metadata, "error": None}
        except Exception as exc:
            return {"content": None, "metadata": None, "error": str(exc)}

    @render.text
    def status() -> str:
        file_info = uploaded_file()
        if file_info is None:
            return "Upload a SoftMouse .xlsx file to begin."

        result = generation_result()
        if result is None:
            return "Upload a SoftMouse .xlsx file to begin."

        if result["error"]:
            return f"Could not build cards: {result['error']}"

        metadata = result["metadata"]
        assert isinstance(metadata, dict)

        lines = [
            f"Ready: {metadata['num_cards']} selected card(s), about {metadata['num_pages']} page(s).",
            f"Selected rows: {metadata.get('num_selected_rows', metadata['num_cards'])} of {metadata.get('num_uploaded_rows', metadata['num_cards'])}",
            f"Comments included: {'yes' if metadata['include_comments'] else 'no'}",
            "Using edited table values.",
        ]

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
