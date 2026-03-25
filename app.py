from __future__ import annotations

import pandas as pd
from shiny import App, Inputs, Outputs, Session, reactive, render, req, ui
from shiny.types import FileInfo

from notecard import build_notecards_bytes, normalize_settings


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
        ui.input_text("protocol_num", "Protocol number", ""),
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
        "Upload a SoftMouse workbook, fill in the contact fields, and download a print-ready notecards.xlsx file."
    ),
    ui.h4("Status"),
    ui.output_text_verbatim("status"),
    ui.h4("Preview of uploaded sheet"),
    ui.output_table("preview"),
    title="Mouse cage cards",
)


def server(input: Inputs, output: Outputs, session: Session):
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
                "contact_name": input.contact_name(),
                "contact_phone": input.contact_phone(),
                "species": input.species(),
            }
        )

    @reactive.calc
    def generation_result() -> dict[str, object] | None:
        file_info = uploaded_file()
        if file_info is None:
            return None

        try:
            content, metadata = build_notecards_bytes(
                xlsx_source=file_info["datapath"],
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
        ]
        warnings_list = metadata.get("warnings", [])
        if warnings_list:
            lines.append("Warnings:")
            lines.extend(f"- {item}" for item in warnings_list)
        return "\n".join(lines)

    @render.table
    def preview():
        file_info = uploaded_file()
        if file_info is None:
            return pd.DataFrame()
        try:
            return pd.read_excel(file_info["datapath"]).head(12)
        except Exception as exc:
            return pd.DataFrame({"error": [str(exc)]})

    @render.download(filename="notecards.xlsx")
    def download_cards():
        result = generation_result()
        req(result is not None)
        req(result["error"] is None)
        content = result["content"]
        req(isinstance(content, (bytes, bytearray)))
        yield content


app = App(app_ui, server)
