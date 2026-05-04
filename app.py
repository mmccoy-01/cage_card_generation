from __future__ import annotations

import pandas as pd
from shiny import App, Inputs, Outputs, Session, reactive, render, req, ui
from shiny.types import FileInfo

from notecard import build_notecards_bytes, normalize_settings


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
    Add editable card-specific columns if the uploaded workbook does not
    already contain an equivalent column.
    """
    df = df.copy()
    existing = {_normalized_column_name(col) for col in df.columns}

    for canonical_name, aliases in EDITABLE_EXTRA_COLUMNS.items():
        if not any(alias in existing for alias in aliases):
            df[canonical_name] = ""

    return df.fillna("")


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
        ui.input_checkbox("include_comments", "Include comments on cards", value=True),
        ui.hr(),
        ui.download_button("download_cards", "Download notecards.xlsx", class_="btn-primary"),
        width=360,
    ),
    ui.h2("Mouse cage card generator"),
    ui.p(
        "Upload a SoftMouse workbook, edit the table if needed, and download a print-ready notecards.xlsx file."
    ),
    ui.h4("Status"),
    ui.output_text_verbatim("status"),
    ui.h4("Editable uploaded sheet"),
    ui.p(
        "You can edit Protocol, Approved, Expires, Source, or any other cell before downloading. "
        "Blank per-cage values fall back to the defaults in the sidebar."
    ),
    ui.output_data_frame("editable_sheet"),
    title="Mouse cage cards",
)


def server(input: Inputs, output: Outputs, session: Session):
    editable_df: reactive.Value[pd.DataFrame] = reactive.Value(pd.DataFrame())

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
        except Exception as exc:
            editable_df.set(pd.DataFrame({"error": [str(exc)]}))

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
        df = editable_df().copy()

        if df.empty or "error" in df.columns:
            return patch["value"]

        value = "" if patch["value"] is None else patch["value"]
        df.iat[patch["row_index"], patch["column_index"]] = value
        editable_df.set(df)

        return value

    @reactive.calc
    def generation_result() -> dict[str, object] | None:
        file_info = uploaded_file()
        if file_info is None:
            return None

        df = editable_df()
        if df.empty:
            return None

        if "error" in df.columns:
            return {"content": None, "metadata": None, "error": str(df["error"].iloc[0])}

        try:
            content, metadata = build_notecards_bytes(
                xlsx_source=df,
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
            return "Upload a SoftMouse .xlsx file to begin."

        result = generation_result()
        if result is None:
            return "Upload a SoftMouse .xlsx file to begin."

        if result["error"]:
            return f"Could not build cards: {result['error']}"

        metadata = result["metadata"]
        assert isinstance(metadata, dict)

        lines = [
            f"Ready: {metadata['num_cards']} card(s), about {metadata['num_pages']} page(s).",
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
