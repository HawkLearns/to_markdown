#!/usr/bin/env python3
"""Reliable PDF, Excel, and Visio to Markdown converter with a minimal dark GUI.

Requirements:
    pip install -U pymupdf4llm pymupdf pandas openpyxl

This script focuses on correctness and stable extraction order by:
- using PyMuPDF4LLM's reading-order aware Markdown conversion in a single pass,
- converting structured Excel requirement exports (Codebeamer/DOORs) to Markdown,
- parsing Visio (.vsdx) files natively into directional Mermaid graph markdown,
- embedding images directly into Markdown as Base64,
- exposing table strategy control,
- appending YAML front matter and TOC metadata,
- preserving page boundaries with visible Markdown dividers.
"""

from __future__ import annotations

import re
import threading
import traceback
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pymupdf
    import pymupdf4llm
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "Missing dependency. Install with: pip install -U pymupdf4llm pymupdf pandas openpyxl"
    ) from exc

try:
    import pandas as pd
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "Missing dependency for Excel support. Install with: pip install -U pandas openpyxl"
    ) from exc

ProgressCallback = Callable[[int, int, str], None]

@dataclass
class ConversionOptions:
    """Runtime options for conversion."""

    split_wide_pages: bool
    table_strategy: str
    embed_images: bool
    force_text: bool


class VisioToMarkdownConverter:
    """Convert a Visio (.vsdx) diagram to Mermaid Markdown."""

    def __init__(self, visio_path: Path, markdown_path: Path) -> None:
        self.visio_path = visio_path
        self.markdown_path = markdown_path

    def convert(self, progress_callback: Optional[ProgressCallback] = None) -> Path:
        """Convert the Visio file and save Markdown output.
       
        Args:
            progress_callback: Optional callback to report conversion progress.

        Returns:
            Path: The path to the generated Markdown file.
           
        Raises:
            ValueError: If the input is a legacy .vsd binary.
            RuntimeError: If parsing the .vsdx archive fails.
        """
        if self.visio_path.suffix.lower() == ".vsd":
            raise ValueError(
                "Legacy .vsd binary format is unsupported. Please convert your file to .vsdx format."
            )

        if progress_callback:
            progress_callback(5, 100, "Opening Visio package...")

        yaml_front_matter = (
            "---\n"
            f"Source: {self.visio_path.name}\n"
            "Parser: Internal VSDX\n"
            "Extraction Type: Mermaid Flowchart\n"
            "---\n\n"
        )
        title = f"## {self.visio_path.stem} Diagrams\n\n"
        page_fragments: list[str] = [yaml_front_matter, title]

        try:
            with zipfile.ZipFile(self.visio_path, "r") as z:
                page_files = [f for f in z.namelist() if f.startswith("visio/pages/page") and f.endswith(".xml")]
                page_files.sort()
               
                total_pages = len(page_files)
                for index, page_file in enumerate(page_files):
                    page_idx = index + 1
                    xml_data = z.read(page_file)
                   
                    mermaid_block = self._process_page(xml_data, page_file)
                    page_fragments.append(mermaid_block)

                    if progress_callback and total_pages > 0:
                        progress = 10 + int((page_idx / total_pages) * 80)
                        progress_callback(progress, 100, f"Processed diagram {page_idx} of {total_pages}...")

        except zipfile.BadZipFile as exc:
            raise RuntimeError(f"Corrupt or invalid .vsdx file structure: {exc}") from exc
        except Exception as exc:
            raise RuntimeError(f"Failed to process Visio elements: {exc}") from exc

        if progress_callback:
            progress_callback(95, 100, "Writing output to disk...")

        final_markdown = "\n".join(page_fragments)
        self.markdown_path.write_text(final_markdown, encoding="utf-8")

        if progress_callback:
            progress_callback(100, 100, "Conversion complete.")

        return self.markdown_path

    def _process_page(self, xml_data: bytes, page_name: str) -> str:
        """Extract shapes and connections from a Visio page XML to Mermaid syntax.

        Args:
            xml_data: The raw XML bytes of the Visio page.
            page_name: The internal name of the page file.

        Returns:
            str: A Mermaid graph Markdown block.
        """
        root = ET.fromstring(xml_data)
       
        for elem in root.iter():
            if "}" in elem.tag:
                elem.tag = elem.tag.split("}", 1)[1]

        shapes: dict[str, str] = {}
        for shape in root.findall(".//Shape"):
            shape_id = shape.attrib.get("ID")
            text_elem = shape.find(".//Text")
           
            if text_elem is not None:
                text = "".join(text_elem.itertext()).strip()
            else:
                text = ""
               
            if shape_id:
                shapes[shape_id] = text

        connects_by_connector: dict[str, dict[str, Optional[str]]] = {}
        for connect in root.findall(".//Connect"):
            from_sheet = connect.attrib.get("FromSheet")
            to_sheet = connect.attrib.get("ToSheet")
            from_cell = connect.attrib.get("FromCell")

            if from_sheet and to_sheet and from_cell:
                if from_sheet not in connects_by_connector:
                    connects_by_connector[from_sheet] = {"begin": None, "end": None}
               
                if "Begin" in from_cell:
                    connects_by_connector[from_sheet]["begin"] = to_sheet
                elif "End" in from_cell:
                    connects_by_connector[from_sheet]["end"] = to_sheet

        fragment = f"### Diagram: {Path(page_name).stem}\n\n```mermaid\ngraph TD\n"
       
        for sid, txt in shapes.items():
            if sid in connects_by_connector and not txt:
                continue
               
            clean_txt = txt.replace("\n", " ").replace('"', "'").strip()
            if not clean_txt:
                clean_txt = f"Node {sid}"
            fragment += f"    {sid}[\"{clean_txt}\"]\n"

        fragment += "\n"
        for cid, ends in connects_by_connector.items():
            if ends["begin"] and ends["end"]:
                fragment += f"    {ends['begin']} --> {ends['end']}\n"

        fragment += "```\n\n---\n"
        return fragment


class ExcelToMarkdownConverter:
    """Convert an Excel requirements export to Markdown."""

    def __init__(self, excel_path: Path, markdown_path: Path) -> None:
        self.excel_path = excel_path
        self.markdown_path = markdown_path

    def convert(self, progress_callback: Optional[ProgressCallback] = None) -> Path:
        """Convert the Excel file and save Markdown output.
       
        Args:
            progress_callback: Optional callback to report conversion progress.

        Returns:
            Path: The path to the generated Markdown file.
        """
        if progress_callback:
            progress_callback(5, 100, "Loading Excel file into pandas...")

        try:
            sheets = pd.read_excel(self.excel_path, sheet_name=None)
        except Exception as exc:
            raise RuntimeError(f"Failed to read Excel file: {exc}") from exc

        if progress_callback:
            progress_callback(20, 100, "Processing Excel rows into Markdown...")

        yaml_front_matter = (
            "---\n"
            f"Source: {self.excel_path.name}\n"
            "Parser: Pandas\n"
            "Extraction Type: Requirements Export\n"
            "---\n\n"
        )

        title = f"## {self.excel_path.stem} Requirements\n\n"
        page_fragments: list[str] = [yaml_front_matter, title]

        total_rows = sum(len(df) for df in sheets.values())
        processed_rows = 0

        for sheet_name, df in sheets.items():
            page_fragments.append(f"### Sheet: {sheet_name}\n\n")

            for index, row in df.iterrows():
                item_idx = index + 1
                fragment = f"### Item {item_idx}\n\n"

                for col in df.columns:
                    val = row[col]
                    if pd.notna(val):
                        clean_val = str(val).replace("\n", "<br>").strip()
                        fragment += f"- **{col}:** {clean_val}\n"

                fragment += "\n---\n"
                page_fragments.append(fragment)
                processed_rows += 1

                if (
                    progress_callback
                    and total_rows > 0
                    and processed_rows % max(1, total_rows // 20) == 0
                ):
                    progress = 20 + int((processed_rows / total_rows) * 70)
                    progress_callback(
                        progress,
                        100,
                        f"Processing row {processed_rows} of {total_rows}...",
                    )

        if progress_callback:
            progress_callback(95, 100, "Writing output to disk...")

        final_markdown = "\n".join(page_fragments)
        self.markdown_path.write_text(final_markdown, encoding="utf-8")

        if progress_callback:
            progress_callback(100, 100, "Conversion complete.")

        return self.markdown_path


class PdfToMarkdownConverter:
    """Convert a PDF to Markdown with full-document pipeline extraction."""

    def __init__(self, pdf_path: Path, markdown_path: Path, options: ConversionOptions) -> None:
        self.pdf_path = pdf_path
        self.markdown_path = markdown_path
        self.options = options

    def convert(self, progress_callback: Optional[ProgressCallback] = None) -> Path:
        """Convert the PDF and save Markdown output.

        Args:
            progress_callback: Optional callback to report conversion progress.

        Returns:
            Path: The path to the generated Markdown file.
        """
        if progress_callback:
            progress_callback(0, 100, "Initializing full-document extraction...")

        yaml_front_matter = (
            "---\n"
            f"Source: {self.pdf_path.name}\n"
            "Parser: PyMuPDF4LLM\n"
            "Extraction Type: Full Document\n"
            "---\n\n"
        )

        doc = pymupdf.open(self.pdf_path)
        toc = doc.get_toc()
        doc.close()

        toc_md = "## Table of Contents\n\n"
        if toc:
            for item in toc:
                level, title, page = item
                indent = "  " * (level - 1)
                toc_md += f"{indent}* {title} (Page {page})\n"
            toc_md += "\n"
        else:
            toc_md += "*(No Table of Contents found)*\n\n"

        if progress_callback:
            progress_callback(20, 100, "Running PyMuPDF4LLM extraction (this may take a while)...")

        try:
            chunks = pymupdf4llm.to_markdown(
                str(self.pdf_path),
                page_chunks=True,
                embed_images=self.options.embed_images,
                force_text=self.options.force_text,
                table_strategy=self.options.table_strategy,
                show_progress=False
            )
        except TypeError:
            chunks = pymupdf4llm.to_markdown(
                str(self.pdf_path),
                page_chunks=True,
                force_text=self.options.force_text,
                show_progress=False
            )

        if progress_callback:
            progress_callback(80, 100, "Assembling document chunks...")

        page_fragments: list[str] = []
        for i, chunk in enumerate(chunks):
            page_number = i + 1
            text = chunk.get("text", "").strip()
            text = self._normalize_markdown(text)
           
            if not text:
                text = "*(No extractable text found on this page.)*"
           
            header = f"---\n### [Document Section: Page {page_number}]\n---\n\n"
            page_fragments.append(header + text)

        final_markdown = yaml_front_matter + toc_md + "\n\n".join(page_fragments).strip() + "\n"

        if progress_callback:
            progress_callback(95, 100, "Writing output to disk...")

        self.markdown_path.write_text(final_markdown, encoding="utf-8")
       
        if progress_callback:
            progress_callback(100, 100, "Conversion complete.")

        return self.markdown_path

    @staticmethod
    def _normalize_markdown(text: str) -> str:
        """Normalize line endings and remove excess newlines."""
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()


class App:
    """Dark, minimal Tkinter GUI."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Document to Markdown Converter")
        self.root.geometry("820x540")
        self.root.minsize(720, 500)

        self.input_files_var = tk.StringVar()
        self.input_paths: tuple[Path, ...] = ()
        self.status_var = tk.StringVar(value="Select PDF, Excel, or Visio files to start.")
        self.progress_var = tk.DoubleVar(value=0.0)

        self.split_wide_pages_var = tk.BooleanVar(value=True)
        self.write_images_var = tk.BooleanVar(value=True)
        self.force_text_var = tk.BooleanVar(value=True)
        self.table_strategy_var = tk.StringVar(value="lines")

        self.colors = {
            "bg": "#0f1115",
            "panel": "#151922",
            "panel_alt": "#1b2130",
            "text": "#e7eaf0",
            "muted": "#9aa3b2",
            "accent": "#7cc0ff",
            "border": "#2a3242",
        }

        self._setup_theme()
        self._build_ui()

    def _setup_theme(self) -> None:
        self.root.configure(bg=self.colors["bg"])
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TFrame", background=self.colors["bg"])
        style.configure("Card.TFrame", background=self.colors["panel"])
        style.configure(
            "TLabel",
            background=self.colors["bg"],
            foreground=self.colors["text"],
            font=("Segoe UI", 10),
        )
        style.configure(
            "Title.TLabel",
            background=self.colors["bg"],
            foreground=self.colors["text"],
            font=("Segoe UI Semibold", 18),
        )
        style.configure(
            "Muted.TLabel",
            background=self.colors["bg"],
            foreground=self.colors["muted"],
            font=("Segoe UI", 10),
        )
        style.configure(
            "TButton",
            padding=(14, 10),
            font=("Segoe UI", 10),
            background=self.colors["panel_alt"],
            foreground=self.colors["text"],
            borderwidth=0,
        )
        style.map(
            "TButton",
            background=[("active", self.colors["accent"]), ("disabled", self.colors["panel_alt"])],
            foreground=[("active", self.colors["bg"]), ("disabled", self.colors["muted"])],
        )
        style.configure(
            "Horizontal.TProgressbar",
            troughcolor=self.colors["panel_alt"],
            background=self.colors["accent"],
            lightcolor=self.colors["accent"],
            darkcolor=self.colors["accent"],
            thickness=10,
        )
        style.configure(
            "Dark.TEntry",
            fieldbackground=self.colors["panel_alt"],
            foreground=self.colors["text"],
            bordercolor=self.colors["border"],
            insertcolor=self.colors["text"],
            padding=8,
        )
        style.configure(
            "Dark.TCheckbutton",
            background=self.colors["panel"],
            foreground=self.colors["text"],
            font=("Segoe UI", 10),
        )
        style.map(
            "Dark.TCheckbutton",
            background=[("active", self.colors["panel"])],
            foreground=[("disabled", self.colors["muted"])],
        )
        style.configure(
            "Dark.TCombobox",
            fieldbackground=self.colors["panel_alt"],
            background=self.colors["panel_alt"],
            foreground=self.colors["text"],
            arrowcolor=self.colors["text"],
            bordercolor=self.colors["border"],
            padding=6,
        )

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=24)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Document to Markdown", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="Reliable extraction pipeline for LLM processing.",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(6, 18))

        card = ttk.Frame(outer, style="Card.TFrame", padding=18)
        card.pack(fill="both", expand=True)

        self._file_picker(
            card,
            "Input files (PDF/Excel/Visio)",
            self.input_files_var,
            self.select_pdf,
            state="readonly",
        )
        ttk.Label(
            card,
            text="Markdown files are saved beside each input with the same filename and a .md extension.",
            style="Muted.TLabel",
            wraplength=730,
        ).pack(anchor="w", pady=(0, 8))

        options_frame = ttk.Frame(card, style="Card.TFrame")
        options_frame.pack(fill="x", pady=(4, 12))

        ttk.Checkbutton(
            options_frame,
            text="Split very wide pages as left then right",
            variable=self.split_wide_pages_var,
            style="Dark.TCheckbutton",
            state="disabled"
        ).grid(row=0, column=0, sticky="w", padx=(0, 18), pady=(0, 10))

        ttk.Checkbutton(
            options_frame,
            text="Embed images as Base64",
            variable=self.write_images_var,
            style="Dark.TCheckbutton",
        ).grid(row=0, column=1, sticky="w", pady=(0, 10))

        ttk.Checkbutton(
            options_frame,
            text="Force text near graphics when possible",
            variable=self.force_text_var,
            style="Dark.TCheckbutton",
        ).grid(row=1, column=0, sticky="w")

        right = ttk.Frame(options_frame, style="Card.TFrame")
        right.grid(row=1, column=1, sticky="w")
        ttk.Label(right, text="Table strategy", style="Muted.TLabel").pack(anchor="w")
        combo = ttk.Combobox(
            right,
            textvariable=self.table_strategy_var,
            values=("lines_strict", "lines", "text"),
            state="readonly",
            style="Dark.TCombobox",
            width=18,
        )
        combo.pack(anchor="w", pady=(4, 0))

        actions = ttk.Frame(card)
        actions.pack(fill="x", pady=(14, 14))

        self.convert_button = ttk.Button(actions, text="Convert", command=self.start_conversion)
        self.convert_button.pack(side="left")
        ttk.Button(actions, text="Exit", command=self.root.destroy).pack(side="right")

        self.progress = ttk.Progressbar(
            card,
            style="Horizontal.TProgressbar",
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
        )
        self.progress.pack(fill="x", pady=(0, 12))

        console = tk.Frame(
            card,
            bg=self.colors["panel_alt"],
            highlightthickness=1,
            highlightbackground=self.colors["border"],
        )
        console.pack(fill="both", expand=True)

        self.status_label = tk.Label(
            console,
            textvariable=self.status_var,
            justify="left",
            anchor="nw",
            bg=self.colors["panel_alt"],
            fg=self.colors["text"],
            font=("Consolas", 10),
            padx=12,
            pady=12,
            wraplength=730,
        )
        self.status_label.pack(fill="both", expand=True)

    def _file_picker(
        self,
        parent: ttk.Frame,
        label: str,
        variable: tk.StringVar,
        callback: Callable[[], None],
        state: str = "normal",
    ) -> None:
        ttk.Label(parent, text=label).pack(anchor="w", pady=(4, 6))
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(0, 12))
        ttk.Entry(row, textvariable=variable, style="Dark.TEntry", state=state).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(row, text="Browse", command=callback).pack(side="left", padx=(10, 0))

    def select_pdf(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Select Input Files",
            filetypes=[
                ("Supported files", "*.pdf *.xlsx *.xls *.vsdx *.vsd"),
                ("PDF files", "*.pdf"),
                ("Excel files", "*.xlsx *.xls"),
                ("Visio files", "*.vsdx *.vsd"),
            ]
        )
        if not paths:
            return
        self.input_paths = tuple(Path(path) for path in paths)
        selected_names = "; ".join(path.name for path in self.input_paths)
        self.input_files_var.set(selected_names)
        self.status_var.set(
            f"{len(self.input_paths)} file(s) selected. Click Convert to start."
        )

    def start_conversion(self) -> None:
        input_paths = self.input_paths
        if not input_paths:
            messagebox.showerror("Invalid Input", "Please select at least one PDF, Excel, or Visio file.")
            return

        supported_extensions = {".pdf", ".xlsx", ".xls", ".vsdx", ".vsd"}
        invalid_paths = [
            path for path in input_paths
            if not path.is_file() or path.suffix.lower() not in supported_extensions
        ]
        if invalid_paths:
            messagebox.showerror(
                "Invalid Input",
                "Please select valid PDF, Excel, or Visio files:\n"
                + "\n".join(str(path) for path in invalid_paths),
            )
            return

        output_paths = tuple(path.with_suffix(".md") for path in input_paths)
        if len(set(output_paths)) != len(output_paths):
            messagebox.showerror(
                "Conflicting Output Names",
                "Selected files in the same folder would produce the same Markdown filename.",
            )
            return

        existing_outputs = [path for path in output_paths if path.exists()]
        if existing_outputs and not messagebox.askyesno(
            "Overwrite Markdown files?",
            "The following output file(s) already exist and will be overwritten:\n\n"
            + "\n".join(str(path) for path in existing_outputs),
        ):
            return

        self.convert_button.state(["disabled"])
        self.progress_var.set(0)
        self.status_var.set(f"Starting conversion of {len(input_paths)} file(s)...")

        options = ConversionOptions(
            split_wide_pages=self.split_wide_pages_var.get(),
            table_strategy=self.table_strategy_var.get(),
            embed_images=self.write_images_var.get(),
            force_text=self.force_text_var.get(),
        )

        worker = threading.Thread(
            target=self._convert_worker,
            args=(input_paths, options),
            daemon=True,
        )
        worker.start()

    def _convert_worker(
        self,
        input_paths: tuple[Path, ...],
        options: ConversionOptions,
    ) -> None:
        try:
            results = []
            total_files = len(input_paths)
            for index, input_path in enumerate(input_paths):
                suffix = input_path.suffix.lower()
                output_path = input_path.with_suffix(".md")
                if suffix in [".xlsx", ".xls"]:
                    converter = ExcelToMarkdownConverter(input_path, output_path)
                elif suffix in [".vsdx", ".vsd"]:
                    converter = VisioToMarkdownConverter(input_path, output_path)
                else:
                    converter = PdfToMarkdownConverter(input_path, output_path, options)

                progress_callback = lambda current, total, message: self._threadsafe_progress(
                    current,
                    total,
                    message,
                    file_index=index,
                    file_count=total_files,
                    file_name=input_path.name,
                )
                results.append(converter.convert(progress_callback=progress_callback))
            self.root.after(0, self._on_success, results)
        except Exception as exc:
            details = "".join(traceback.format_exception_only(type(exc), exc)).strip()
            self.root.after(
                0,
                self._on_failure,
                f"{input_path.name}: {details}",
            )

    def _threadsafe_progress(
        self,
        current: int,
        total: int,
        message: str,
        file_index: int = 0,
        file_count: int = 1,
        file_name: str = "",
    ) -> None:
        total = max(total, 1)
        file_count = max(file_count, 1)
        file_percent = min(max(current / total, 0.0), 1.0)
        percent = ((file_index + file_percent) / file_count) * 100.0
        status = f"File {file_index + 1} of {file_count}"
        if file_name:
            status += f" ({file_name})"
        self.root.after(0, self._update_progress, percent, f"{status}: {message}")

    def _update_progress(self, percent: float, message: str) -> None:
        """Update the progress bar widget and status label text thread-safely.

        Args:
            percent: The percentage completion value.
            message: The descriptive status message.
        """
        self.status_var.set(message)
        self.progress_var.set(percent)

    def _on_success(self, result_paths: list[Path]) -> None:
        self.convert_button.state(["!disabled"])
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress_var.set(100)
        results = "\n".join(str(path) for path in result_paths)
        self.status_var.set(
            f"Conversion finished successfully for {len(result_paths)} file(s).\n\n"
            f"Markdown files:\n{results}"
        )
        messagebox.showinfo(
            "Done",
            f"Converted {len(result_paths)} file(s). Markdown saved to:\n{results}",
        )

    def _on_failure(self, details: str) -> None:
        self.convert_button.state(["!disabled"])
        self.progress.stop()
        self.progress.configure(mode="determinate")
        self.progress_var.set(0)
        self.status_var.set(f"Conversion failed.\n\n{details}")
        messagebox.showerror("Conversion failed", details)


def main() -> None:
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
