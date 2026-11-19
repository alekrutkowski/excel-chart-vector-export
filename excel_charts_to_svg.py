#!/usr/bin/env python3
"""Export embedded Excel charts and chart sheets to inspected SVG and/or EMF files.

Windows: native Excel COM -> one-chart PDF -> SVG via PyMuPDF;
         chart picture clipboard -> native EMF (with optional Inkscape fallback).
Windows/Linux: optional licensed Aspose.Cells for Python via Java -> SVG/EMF.

The source workbook is always opened read-only and is never saved.
"""
from __future__ import annotations

import argparse
from collections import Counter
from contextlib import contextmanager, suppress
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from hashlib import sha256
import json
import os
from pathlib import Path
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import time
import unicodedata
from typing import Any, Iterator
from uuid import uuid4
import xml.etree.ElementTree as ET
from zipfile import BadZipFile, ZipFile

VERSION = "1.2.1"
EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls", ".xltx", ".xltm", ".xlt"}
SVG_NS = "http://www.w3.org/2000/svg"
XL_WORKSHEET = -4167
XL_CHART = -4109
XL_LOCATION_AS_NEW_SHEET = 1
XL_VISIBLE = -1
XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0
XL_SCREEN = 1
XL_PRINTER = 2
XL_PICTURE = -4147
CF_ENHMETAFILE = 14

# Windows GDI Enhanced Metafile (EMF) record types which paint actual vector
# geometry or text. Some charts combine these with embedded raster records.
EMF_VECTOR_RECORDS = frozenset({
    2, 3, 4, 5, 6, 7, 8, 41, 42, 43, 44, 45, 46, 47, 54, 55, 56,
    62, 63, 64, 71, 72, 74, 83, 84, 85, 86, 87, 88, 89, 90, 91,
    92, 96, 97,
})
EMF_RASTER_RECORDS = frozenset({15, 76, 77, 78, 79, 80, 81, 114, 116})


class ExtractionError(RuntimeError):
    """A chart was not rendered or its SVG did not pass validation."""


@dataclass
class ChartSpec:
    sheet: str
    object_name: str
    title: str
    position: int
    chart: Any = field(repr=False)
    chart_type: str = ""
    name_hint: str = ""
    host_sheet: Any = field(default=None, repr=False)
    chart_object: Any = field(default=None, repr=False)


@dataclass
class SVGInfo:
    width: str
    height: str
    vector_elements: int
    raster_elements: int
    external_images: int


@dataclass
class EMFInfo:
    """EMF structural and GDI-record summary, not a visual-fidelity guarantee."""
    records: int
    vector_elements: int
    raster_elements: int
    emf_plus_comments: int


@dataclass
class Result:
    source: str
    sheet: str
    object_name: str
    title: str
    chart_type: str
    output: str | None
    status: str
    format: str = "svg"
    render_method: str | None = None
    vector_elements: int | None = None
    raster_elements: int | None = None
    error: str | None = None


def clean_filename(value: str, maximum: int = 80) -> str:
    """Make a readable filename that is valid on Windows and Linux."""
    text = unicodedata.normalize("NFC", str(value or ""))
    text = re.sub(r'[\\/:*?"<>|\x00-\x1f\x7f]', "-", text)
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"-{2,}", "-", text).strip(" .-_")
    if len(text) > maximum:
        text = text[:maximum].rstrip(" .-_")
    if not text:
        text = "Untitled"
    # Windows device names, even with an extension, are disallowed.
    if re.fullmatch(r"(?i:con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\..*)?", text):
        text = "_" + text
    return text


def chart_label(title: str, object_name: str, position: int, name_hint: str = "") -> str:
    title = (title or "").strip()
    if title:
        return title
    object_name = (object_name or "").strip()
    generic_object = bool(re.fullmatch(r"(?i:chart)(?:\s*\d+)?", object_name))
    if object_name and not generic_object:
        return object_name
    if name_hint:
        return name_hint
    return f"Chart {position:02d}" if generic_object or not object_name else object_name


def descriptive_series_names(raw_names: list[str]) -> str:
    """Conservatively derive a human-friendly filename for an untitled chart."""
    names = []
    for item in raw_names:
        name = str(item or "").strip().strip('"')
        if (not name or len(name) > 72 or name.startswith("=")
                or re.fullmatch(r"(?i:series|reihe|série)\s*\d+", name)):
            continue
        if name.casefold() not in [x.casefold() for x in names]:
            names.append(name)
    return " and ".join(names[:2])


def excel_series_hint(chart: Any) -> str:
    try:
        collection = chart.SeriesCollection()
        names = [str(collection(i).Name) for i in range(1, min(2, collection.Count) + 1)]
        return descriptive_series_names(names)
    except Exception:
        return ""


def aspose_series_hint(chart: Any) -> str:
    try:
        collection = chart.getNSeries()
        names = [str(collection.get(i).getName())
                 for i in range(min(2, collection.getCount()))]
        return descriptive_series_names(names)
    except Exception:
        return ""


def collect_inputs(expressions: list[str], recursive: bool) -> list[Path]:
    """Accept paths, directories and unexpanded shell globs."""
    import glob

    paths: set[Path] = set()
    for expr in expressions:
        raw = Path(expr).expanduser()
        if raw.is_file():
            candidates = [raw]
        elif raw.is_dir():
            candidates = raw.rglob("*") if recursive else raw.iterdir()
        else:
            candidates = (Path(x) for x in glob.glob(str(raw), recursive=True))
        for candidate in candidates:
            if (candidate.is_file() and candidate.suffix.lower() in EXTENSIONS
                    and not candidate.name.startswith("~$")):
                paths.add(candidate.resolve())
    return sorted(paths, key=lambda p: str(p).casefold())


def count_ooxml_chart_parts(path: Path) -> int | None:
    """Audit charts reachable from OOXML relationships, when those exist.

    Count chart references in relationship parts, including new chartEx parts.
    Fall back to counting chart XML parts for incomplete or unusual packages.
    An ordinary legacy .xls workbook is not a ZIP and returns None.
    """
    if not __import__("zipfile").is_zipfile(path):
        return None
    with ZipFile(path) as archive:
        names = archive.namelist()
        chart_parts = [p for p in names
                       if re.fullmatch(r"xl/charts/chart(?:ex)?\d+\.xml", p.lower())]
        rels_parts = [p for p in names if p.lower().startswith("xl/")
                      and p.lower().endswith(".rels")]
        references = 0
        for rels_name in rels_parts:
            root = ET.fromstring(archive.read(rels_name))
            for item in root:
                kind = str(item.get("Type", "")).rstrip("/").rsplit("/", 1)[-1].lower()
                if kind in {"chart", "chartex"} and item.get("TargetMode") != "External":
                    references += 1
        # A referenced chart has a real on-sheet representation; an orphaned
        # chart part can exist after a chart was deleted in Excel.
        return references if rels_parts else len(chart_parts)


def safe_attr(obj: Any, attr: str, fallback: Any = "") -> Any:
    try:
        result = getattr(obj, attr)
        return result if result is not None else fallback
    except Exception:
        return fallback


def excel_title(chart: Any) -> str:
    try:
        if chart.HasTitle:
            return str(chart.ChartTitle.Text).strip()
    except Exception:
        pass
    return ""


def aspose_title(chart: Any) -> str:
    try:
        title = chart.getTitle()
        return str(title.getText() or "").strip()
    except Exception:
        return ""


def group_chart_shapes(shapes: Any) -> Iterator[Any]:
    """Find chart-bearing shapes, including charts inside shape groups."""
    for shape in shapes:
        if safe_attr(shape, "Type", None) == 6:  # msoGroup
            items = safe_attr(shape, "GroupItems", None)
            if items is not None:
                for index in range(1, int(items.Count) + 1):
                    try:
                        child = items.Item(index)
                    except Exception:
                        child = items(index)
                    yield from group_chart_shapes([child])
        elif safe_attr(shape, "HasChart", 0) == -1:  # msoTrue
            yield shape


def excel_specs(workbook: Any) -> Iterator[ChartSpec]:
    """Enumerate ChartObjects, additional chart shapes, and chart sheets."""
    for ws in workbook.Worksheets:
        seen: set[str] = set()
        collection = ws.ChartObjects()
        for index in range(1, collection.Count + 1):
            chart_object = collection(index)
            chart = chart_object.Chart
            seen.add(str(chart_object.Name).casefold())
            yield ChartSpec(str(ws.Name), str(chart_object.Name), excel_title(chart),
                            index, chart, str(safe_attr(chart, "ChartType")),
                            name_hint=excel_series_hint(chart), host_sheet=ws, chart_object=chart_object)
        # Some charts are accessible as shapes rather than ChartObjects,
        # especially inside Excel drawing groups. Avoid counting duplicates.
        next_index = collection.Count
        for shape in group_chart_shapes(ws.Shapes):
            name = str(shape.Name)
            if name.casefold() in seen:
                continue
            seen.add(name.casefold())
            next_index += 1
            chart = shape.Chart
            yield ChartSpec(str(ws.Name), name, excel_title(chart), next_index,
                            chart, str(safe_attr(chart, "ChartType")),
                            name_hint=excel_series_hint(chart), host_sheet=ws, chart_object=shape)
    for index, chart in enumerate(workbook.Charts, 1):
        yield ChartSpec(str(chart.Name), str(chart.Name), excel_title(chart),
                        index, chart, str(safe_attr(chart, "ChartType")),
                        name_hint=excel_series_hint(chart), host_sheet=chart)


def aspose_specs(workbook: Any) -> Iterator[ChartSpec]:
    worksheets = workbook.getWorksheets()
    for i in range(worksheets.getCount()):
        ws = worksheets.get(i)
        charts = ws.getCharts()
        for j in range(charts.getCount()):
            chart = charts.get(j)
            type_getter = safe_attr(chart, "getType", None)
            chart_type = str(type_getter()) if callable(type_getter) else ""
            yield ChartSpec(str(ws.getName()), str(chart.getName() or f"Chart {j+1}"),
                            aspose_title(chart), j + 1, chart, chart_type,
                            name_hint=aspose_series_hint(chart))


def unique_output(folder: Path, spec: ChartSpec, seen: set[str], extension: str = "svg") -> Path:
    stem = f"{clean_filename(spec.sheet, 56)}__{clean_filename(chart_label(spec.title, spec.object_name, spec.position, spec.name_hint), 84)}"
    candidate = stem
    serial = 2
    while candidate.casefold() in seen:
        candidate = f"{stem}__{serial:02d}"
        serial += 1
    seen.add(candidate.casefold())
    return folder / f"{candidate}.{extension}"


def svg_inspect_and_label(path: Path, *, title: str, description: str, require_vector: bool) -> SVGInfo:
    """Validate SVG syntax, self-containment and graphics, and add accessibility metadata."""
    ET.register_namespace("", SVG_NS)
    ET.register_namespace("xlink", "http://www.w3.org/1999/xlink")
    try:
        tree = ET.parse(path)
        root = tree.getroot()
    except (ET.ParseError, OSError, ValueError) as exc:
        raise ExtractionError(f"Invalid SVG: {exc}") from exc
    if root.tag != f"{{{SVG_NS}}}svg":
        raise ExtractionError("The renderer did not produce an SVG root element")
    vectors = {"path", "rect", "circle", "ellipse", "line", "polyline", "polygon", "text", "use"}
    vector_count = 0
    raster_count = 0
    external_count = 0
    for element in root.iter():
        local = element.tag.rsplit("}", 1)[-1]
        if local in {"script", "foreignObject"}:
            raise ExtractionError(f"Unsafe SVG element is not allowed: {local}")
        if local in vectors:
            vector_count += 1
        if local == "image":
            raster_count += 1
        # Inspect all SVG link attributes, not just <image>: an external <use>,
        # stylesheet, filter or event handler can make the file unsafe/nonportable.
        for name, raw_value in element.attrib.items():
            attr = name.rsplit("}", 1)[-1].lower()
            value = raw_value.strip()
            if attr.startswith("on"):
                raise ExtractionError(f"Unsafe SVG event handler: {attr}")
            if attr in {"href", "src"} and value and not (
                    value.startswith("#") or value.startswith("data:image/")):
                external_count += 1
            if attr in {"style", "fill", "stroke", "filter", "clip-path", "mask"}:
                refs = re.findall(r"url\(\s*['\"]?([^'\"\)\s]+)", value, flags=re.IGNORECASE)
                external_count += sum(not link.startswith("#") for link in refs)
        if local == "style" and element.text:
            css = element.text
            if re.search(r"@import\b", css, flags=re.IGNORECASE):
                external_count += 1
            refs = re.findall(r"url\(\s*['\"]?([^'\"\)\s]+)", css, flags=re.IGNORECASE)
            external_count += sum(not link.startswith("#") for link in refs)
    if vector_count + raster_count == 0:
        raise ExtractionError("SVG is empty (no graphical elements found)")
    if require_vector and vector_count == 0:
        raise ExtractionError("Chart rendered as a raster image inside SVG (use --allow-raster to permit this)")
    if external_count:
        raise ExtractionError(f"SVG refers to {external_count} external image(s) or resource(s) and is not self-contained")
    for child in list(root):
        if child.tag in (f"{{{SVG_NS}}}title", f"{{{SVG_NS}}}desc"):
            root.remove(child)
    title_node = ET.Element(f"{{{SVG_NS}}}title")
    title_node.text = title
    desc_node = ET.Element(f"{{{SVG_NS}}}desc")
    desc_node.text = description
    root.insert(0, desc_node)
    root.insert(0, title_node)
    root.set("role", "img")
    root.set("aria-label", title)
    tree.write(path, encoding="utf-8", xml_declaration=True)
    return SVGInfo(root.get("width", ""), root.get("height", ""),
                   vector_count, raster_count, external_count)


def emf_inspect(path: Path, *, require_vector: bool = True) -> EMFInfo:
    """Reject renamed images, truncated files, malformed records and raster-only EMFs.

    EMF is a *binary record stream*, not an XML/SVG wrapper. The header must
    begin with EMR_HEADER, include the Microsoft EMF signature at byte 40,
    correctly declare the total byte count and end in EMR_EOF. We inspect the
    records to catch image-only containers; records can still contain raster
    elements or complex EMF+ payloads, which are reported in the manifest.
    """
    raw = path.read_bytes()
    if len(raw) < 108:
        raise ExtractionError("EMF file is too short to contain a valid header and EOF")
    kind, header_size = struct.unpack_from("<II", raw)
    if kind != 1 or header_size < 88 or header_size > len(raw) or header_size % 4:
        raise ExtractionError("Invalid EMF header record")
    if raw[40:44] != b" EMF":
        raise ExtractionError("Not an Enhanced Metafile (missing EMF signature)")
    declared_bytes, declared_records = struct.unpack_from("<II", raw, 48)
    if declared_bytes != len(raw) or declared_records < 2:
        raise ExtractionError("Invalid EMF declared byte length or record count")

    offset = 0
    total = 0
    vectors = 0
    rasters = 0
    emf_plus = 0
    last_type = None
    while offset < len(raw):
        if len(raw) - offset < 8:
            raise ExtractionError("Truncated EMF record")
        rec_type, rec_size = struct.unpack_from("<II", raw, offset)
        if rec_size < 8 or rec_size % 4 or offset + rec_size > len(raw):
            raise ExtractionError(f"Invalid EMF record size at byte {offset}")
        total += 1
        vectors += int(rec_type in EMF_VECTOR_RECORDS)
        rasters += int(rec_type in EMF_RASTER_RECORDS)
        if rec_type == 70 and b"EMF+" in raw[offset + 8:offset + rec_size]:
            emf_plus += 1
        last_type = rec_type
        offset += rec_size
    if total != declared_records or last_type != 14:
        raise ExtractionError("EMF record count mismatch or missing final EOF record")
    if require_vector and vectors == 0:
        raise ExtractionError(
            "No classic GDI vector drawing records found in EMF; possible raster-only "
            "or EMF+-only output (use --allow-raster only if acceptable)")
    return EMFInfo(total, vectors, rasters, emf_plus)


def clipboard_sequence() -> int | None:
    """Best-effort guard against accidentally reading a previous chart's EMF."""
    try:
        import ctypes
        return int(ctypes.windll.user32.GetClipboardSequenceNumber())
    except (AttributeError, OSError):
        return None


def get_emf_clipboard_bytes(*, attempts: int = 6, pause: float = 0.15) -> bytes:
    """Read raw CF_ENHMETAFILE bytes via pywin32 with clipboard-busy retries."""
    try:
        import win32clipboard
    except ImportError as exc:
        raise ExtractionError("EMF clipboard export requires pywin32") from exc
    problem = "Excel did not place CF_ENHMETAFILE on the Windows clipboard"
    for n in range(attempts):
        opened = False
        try:
            win32clipboard.OpenClipboard()
            opened = True
            if not win32clipboard.IsClipboardFormatAvailable(CF_ENHMETAFILE):
                raise ExtractionError(problem)
            data = win32clipboard.GetClipboardData(CF_ENHMETAFILE)
            if not isinstance(data, bytes) or not data:
                raise ExtractionError("Windows clipboard did not return EMF binary data")
            return data
        except Exception as exc:
            problem = str(exc)
            if n + 1 == attempts:
                break
        finally:
            if opened:
                win32clipboard.CloseClipboard()
        time.sleep(pause)
    raise ExtractionError(f"Could not read native Excel EMF: {problem}")


def inkscape_executable(explicit: Path | None) -> Path | None:
    """Discover an optional SVG-to-EMF converter for clipboard failure cases."""
    if explicit is not None:
        return explicit.expanduser().resolve()
    found = shutil.which("inkscape")
    if found:
        return Path(found)
    if sys.platform.startswith("win"):
        for env_name in ("ProgramW6432", "ProgramFiles", "ProgramFiles(x86)"):
            root = os.environ.get(env_name)
            if root:
                candidate = Path(root) / "Inkscape" / "bin" / "inkscape.exe"
                if candidate.is_file():
                    return candidate
    return None


def inkscape_svg_to_emf(svg: Path, emf: Path, executable: Path) -> None:
    """Fallback conversion: SVG vector paths to native binary EMF via Inkscape."""
    command = [str(executable), str(svg), f"--export-filename={emf}"]
    try:
        result = subprocess.run(command, capture_output=True, text=True,
                                timeout=120, check=False)
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise ExtractionError(f"Inkscape SVG-to-EMF conversion failed: {exc}") from exc
    if result.returncode or not emf.is_file() or emf.stat().st_size < 108:
        detail = (result.stderr or result.stdout or "no EMF produced").strip()
        raise ExtractionError(f"Inkscape could not export EMF ({detail})")


def pdf_to_svg(pdf: Path, svg: Path, *, crop: str, keep_text: bool) -> None:
    """Convert a one-page PDF to true vector SVG, with optional ink-bound cropping."""
    try:
        import pymupdf as fitz
    except ImportError as exc:
        raise ExtractionError("PyMuPDF is required for the Excel backend: pip install pymupdf") from exc

    try:
        with fitz.open(pdf) as document:
            if len(document) != 1:
                raise ExtractionError(f"Expected one PDF page per chart; got {len(document)}")
            page = document[0]
            if crop == "auto":
                boxes = []
                for item in page.get_bboxlog():
                    if item[0] == "ignore-text":
                        continue
                    box = fitz.Rect(item[1]) & page.rect
                    if box.is_empty or box.width < .25 or box.height < .25:
                        continue
                    boxes.append(box)
                # Page-sized white backgrounds or printer margins should not
                # force the exported chart to cover the whole sheet of paper.
                remaining = [r for r in boxes if r.get_area() < page.rect.get_area() * .94]
                if remaining:
                    result = remaining[0]
                    for box in remaining[1:]:
                        result |= box
                    if result.get_area() > 16:
                        pad = max(3.0, min(result.width, result.height) * .012)
                        result = (fitz.Rect(result.x0 - pad, result.y0 - pad,
                                            result.x1 + pad, result.y1 + pad) & page.rect)
                        page.set_cropbox(result)
            body = page.get_svg_image(text_as_path=not keep_text)
        svg.write_text(body, encoding="utf-8")
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionError(f"PDF-to-SVG failed: {exc}") from exc


@contextmanager
def temporarily_visible_chart(spec: ChartSpec) -> Iterator[None]:
    """Unhide an original chart only in Excel memory, then restore its state.

    This is important even when Excel reports no exception: printing an
    invisible embedded chart can produce the entire worksheet as a PDF.
    """
    sheet = spec.host_sheet
    obj = spec.chart_object
    sheet_visibility = safe_attr(sheet, "Visible", XL_VISIBLE)
    object_visibility = safe_attr(obj, "Visible", True)
    hidden_sheet = sheet is not None and sheet_visibility != XL_VISIBLE
    hidden_object = obj is not None and not bool(object_visibility)
    try:
        if hidden_sheet:
            sheet.Visible = XL_VISIBLE
        if hidden_object:
            obj.Visible = True
        yield
    finally:
        if hidden_object:
            with suppress(Exception):
                obj.Visible = object_visibility
        if hidden_sheet:
            with suppress(Exception):
                sheet.Visible = sheet_visibility


def check_chart_pdf(pdf: Path) -> None:
    """Reject missing, empty, or multi-page exports before making an SVG."""
    if not pdf.is_file() or pdf.stat().st_size < 100:
        raise ExtractionError("Excel PDF export did not produce a usable file")
    try:
        import pymupdf
        with pymupdf.open(pdf) as document:
            pages = len(document)
    except Exception as exc:
        raise ExtractionError(f"Excel did not produce a readable PDF: {exc}") from exc
    if pages != 1:
        raise ExtractionError(f"Expected a single-chart PDF, but Excel printed {pages} page(s)")


class ExcelBackend:
    name = "excel"

    def __init__(self, crop: str, keep_text: bool, show_excel: bool = False,
                 inkscape_path: Path | None = None, require_vector: bool = True):
        if not sys.platform.startswith("win"):
            raise ExtractionError("Native Excel engine requires Windows and desktop Microsoft Excel")
        try:
            import pythoncom
            import win32com.client
            import pymupdf  # noqa: F401
        except ImportError as exc:
            raise ExtractionError("Excel engine requires pywin32 and pymupdf: pip install pywin32 pymupdf") from exc
        self._pythoncom = pythoncom
        self._client = win32com.client
        self.crop = crop
        self.keep_text = keep_text
        self.require_vector = require_vector
        self.inkscape_path = inkscape_executable(inkscape_path)
        self._pythoncom.CoInitialize()
        try:
            self.app = self._client.DispatchEx("Excel.Application")
            self.app.Visible = show_excel
            self.app.DisplayAlerts = False
            self.app.EnableEvents = False
            self.app.AskToUpdateLinks = False
            self.app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        except Exception as exc:
            with suppress(Exception):
                self.app.Quit()
            self._pythoncom.CoUninitialize()
            raise ExtractionError(f"Could not start Microsoft Excel: {exc}") from exc

    def open(self, path: Path) -> Any:
        try:
            return self.app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True,
                                           IgnoreReadOnlyRecommended=True, AddToMru=False,
                                           Notify=False, Password="")
        except Exception as exc:
            raise ExtractionError(f"Cannot open workbook in Excel: {exc}") from exc

    def enumerate(self, workbook: Any) -> Iterator[ChartSpec]:
        return excel_specs(workbook)

    @staticmethod
    def _export_pdf(chart: Any, pdf: Path) -> None:
        # Do not set From=1, To=1: that would conceal a worksheet export bug
        # by saving only the first page of the incorrect worksheet PDF.
        chart.ExportAsFixedFormat(
            XL_TYPE_PDF, str(pdf), XL_QUALITY_STANDARD, False, True)
        check_chart_pdf(pdf)

    def _activate(self, spec: ChartSpec) -> Any:
        """Select the actual chart, so Excel prints it rather than its sheet.

        An embedded chart's ExportAsFixedFormat *can* silently print its entire
        worksheet unless its ChartObject is active. Excel's ActiveChart is the
        object corresponding to the most recently activated chart.
        """
        if spec.chart_object is None:
            spec.chart.Activate()  # independent chart sheet
        else:
            spec.host_sheet.Activate()
            activate = safe_attr(spec.chart_object, "Activate", None)
            if callable(activate):
                activate()
            else:
                # A chart contained in a grouped Shape may not expose
                # ChartObject.Activate, but Shape.Select can still work.
                select = safe_attr(spec.chart_object, "Select", None)
                if not callable(select):
                    raise ExtractionError("Chart object cannot be activated or selected")
                select()
                spec.chart.ChartArea.Select()
        selected = safe_attr(self.app, "ActiveChart", None)
        if selected is None:
            raise ExtractionError("Excel did not activate the requested chart")
        return selected

    def _export_using_duplicate_chart_sheet(self, spec: ChartSpec, pdf: Path) -> None:
        """Fallback: export a copy of the chart on its own chart sheet.

        Duplicating keeps the user's original chart untouched in Excel memory.
        Only the duplicate is relocated, and the temporary chart sheet is
        deleted. The read-only workbook is always closed without saving.
        """
        if spec.chart_object is None:
            raise ExtractionError("Cannot use the duplicate fallback for a chart sheet")
        duplicate = None
        chart_sheet = None
        try:
            duplicate = spec.chart_object.Duplicate()
            duplicate_chart = safe_attr(duplicate, "Chart", None)
            if duplicate_chart is None:
                raise ExtractionError("Excel duplicated a drawing but not an editable chart")
            # Location destroys the duplicated embedded ChartObject, not the
            # source chart. The 31-character sheet name is safe and unique.
            chart_sheet = duplicate_chart.Location(
                XL_LOCATION_AS_NEW_SHEET, "_svg_" + uuid4().hex[:16])
            chart_sheet.Activate()
            if pdf.exists():
                pdf.unlink()
            self._export_pdf(chart_sheet, pdf)
        finally:
            if chart_sheet is not None:
                with suppress(Exception):
                    chart_sheet.Delete()
            elif duplicate is not None:
                with suppress(Exception):
                    duplicate.Delete()

    def _render_svg(self, spec: ChartSpec, destination: Path) -> str:
        with tempfile.TemporaryDirectory(prefix="chart-svg-", dir=destination.parent) as tmp:
            pdf = Path(tmp) / "chart.pdf"
            with temporarily_visible_chart(spec):
                first_error = None
                try:
                    active_chart = self._activate(spec)
                    self._export_pdf(active_chart, pdf)
                    method = "activated_chart"
                except Exception as exc:
                    first_error = exc
                    if spec.chart_object is None:
                        raise ExtractionError(f"Chart-sheet PDF export failed: {exc}") from exc
                    try:
                        self._export_using_duplicate_chart_sheet(spec, pdf)
                        method = "duplicate_chart_sheet"
                    except Exception as fallback_exc:
                        raise ExtractionError(
                            "Both chart-only PDF methods failed. "
                            f"Activated chart: {first_error}; "
                            f"duplicate chart sheet: {fallback_exc}"
                        ) from fallback_exc
                pdf_to_svg(pdf, destination, crop=self.crop, keep_text=self.keep_text)
            return method

    def _render_native_emf(self, spec: ChartSpec, destination: Path) -> str:
        """Copy the actual chart as a GDI picture and capture CF_ENHMETAFILE.

        On Windows, Excel.ChartObject.CopyPicture / Chart.CopyPicture with
        Format=xlPicture place a genuine Enhanced Metafile on the clipboard.
        This deliberately never disguises a PNG as an EMF. Since Excel uses
        the global clipboard, warn users it will replace its current content.
        """
        with temporarily_visible_chart(spec):
            active_chart = self._activate(spec)
            last_error: Exception | None = None
            for attempt in range(3):
                before = clipboard_sequence()
                try:
                    if spec.chart_object is not None:
                        spec.chart_object.CopyPicture(Appearance=XL_SCREEN,
                                                      Format=XL_PICTURE)
                    else:
                        active_chart.CopyPicture(Appearance=XL_SCREEN,
                                                 Format=XL_PICTURE, Size=XL_PRINTER)
                    # Check the clipboard really changed, instead of silently
                    # exporting the previous chart when Excel's copy failed.
                    if before:
                        changed = False
                        for _ in range(6):
                            after = clipboard_sequence()
                            if after is None or after != before:
                                changed = True
                                break
                            time.sleep(0.1)
                        if not changed:
                            raise ExtractionError("Excel did not update the Windows clipboard")
                    destination.write_bytes(get_emf_clipboard_bytes())
                    emf_inspect(destination, require_vector=self.require_vector)
                    return "excel_clipboard_emf"
                except Exception as exc:
                    last_error = exc
                    destination.unlink(missing_ok=True)
                    if attempt < 2:
                        time.sleep(0.15)
            raise ExtractionError(f"Excel could not copy the chart as vector EMF: {last_error}")

    def _render_emf(self, spec: ChartSpec, destination: Path) -> str:
        try:
            return self._render_native_emf(spec, destination)
        except Exception as native_exc:
            destination.unlink(missing_ok=True)
            if self.inkscape_path is None:
                raise ExtractionError(
                    f"Native Excel EMF export failed: {native_exc}. "
                    "Try --show-excel, or install Inkscape for automatic SVG-to-EMF fallback."
                ) from native_exc
            try:
                with tempfile.TemporaryDirectory(prefix="chart-emf-",
                                                 dir=destination.parent) as tmp:
                    svg = Path(tmp) / "chart.svg"
                    self._render_svg(spec, svg)
                    svg_inspect_and_label(svg, title=chart_label(spec.title, spec.object_name,
                                                                spec.position, spec.name_hint),
                                          description="Intermediate vector chart for EMF conversion",
                                          require_vector=self.require_vector)
                    inkscape_svg_to_emf(svg, destination, self.inkscape_path)
                    emf_inspect(destination, require_vector=self.require_vector)
                return "inkscape_svg_fallback"
            except Exception as fallback_exc:
                destination.unlink(missing_ok=True)
                raise ExtractionError(
                    f"EMF export failed. Native Excel clipboard: {native_exc}; "
                    f"Inkscape fallback: {fallback_exc}") from fallback_exc

    def render(self, spec: ChartSpec, destination: Path) -> str:
        if destination.suffix.lower() == ".svg":
            return self._render_svg(spec, destination)
        if destination.suffix.lower() == ".emf":
            return self._render_emf(spec, destination)
        raise ExtractionError(f"Unsupported output format: {destination.suffix}")

    def close_workbook(self, book: Any) -> None:
        book.Close(SaveChanges=False)

    def close(self) -> None:
        try:
            self.app.Quit()
        finally:
            self._pythoncom.CoUninitialize()


class AsposeBackend:
    name = "aspose"

    def __init__(self, license_path: Path | None):
        try:
            import jpype
            import asposecells
        except ImportError as exc:
            raise ExtractionError("Aspose engine requires: pip install aspose-cells") from exc
        self._jpype = jpype
        self._started_jvm = not jpype.isJVMStarted()
        if self._started_jvm:
            jpype.startJVM()
        try:
            from asposecells.api import Workbook, ImageOrPrintOptions, ImageType, License
            self.Workbook = Workbook
            self.ImageOrPrintOptions = ImageOrPrintOptions
            self.ImageType = ImageType
            if license_path is not None:
                License().setLicense(str(license_path.resolve()))
        except Exception as exc:
            if self._started_jvm:
                jpype.shutdownJVM()
            raise ExtractionError(f"Cannot initialize Aspose.Cells: {exc}") from exc

    def open(self, path: Path) -> Any:
        try:
            return self.Workbook(str(path))
        except Exception as exc:
            raise ExtractionError(f"Cannot open workbook in Aspose.Cells: {exc}") from exc

    def enumerate(self, workbook: Any) -> Iterator[ChartSpec]:
        return aspose_specs(workbook)

    def render(self, spec: ChartSpec, destination: Path) -> str:
        extension = destination.suffix.lower()
        if extension not in {".svg", ".emf"}:
            raise ExtractionError(f"Unsupported Aspose output format: {extension}")
        try:
            options = self.ImageOrPrintOptions()
            options.setImageType(self.ImageType.SVG if extension == ".svg"
                                 else self.ImageType.EMF)
            spec.chart.toImage(str(destination), options)
        except Exception as exc:
            raise ExtractionError(f"Aspose {extension} rendering failed: {exc}") from exc
        if not destination.is_file() or destination.stat().st_size < 100:
            raise ExtractionError(f"Aspose returned an empty {extension} "
                                  "(chart type may be unsupported)")
        return f"aspose_{extension.lstrip('.')}"

    def close_workbook(self, book: Any) -> None:
        pass  # Java garbage collector manages workbook lifetime.

    def close(self) -> None:
        if self._started_jvm:
            self._jpype.shutdownJVM()


def backend_for(engine: str, *, crop: str, keep_text: bool,
                license_path: Path | None, show_excel: bool = False,
                inkscape_path: Path | None = None,
                require_vector: bool = True) -> ExcelBackend | AsposeBackend:
    if engine == "excel":
        return ExcelBackend(crop, keep_text, show_excel=show_excel,
                            inkscape_path=inkscape_path, require_vector=require_vector)
    if engine == "aspose":
        return AsposeBackend(license_path)
    if sys.platform.startswith("win"):
        try:
            return ExcelBackend(crop, keep_text, show_excel=show_excel,
                                inkscape_path=inkscape_path,
                                require_vector=require_vector)
        except ExtractionError as exc:
            print(f"Note: native Excel unavailable ({exc}); trying Aspose.Cells", file=sys.stderr)
    return AsposeBackend(license_path)


def workbook_dirs(inputs: list[Path], output_dir: Path) -> dict[Path, Path]:
    counts = Counter(clean_filename(path.stem, 70).casefold() for path in inputs)
    result = {}
    for path in inputs:
        basename = clean_filename(path.stem, 70)
        if counts[basename.casefold()] > 1:
            suffix = sha256(str(path).encode("utf-8")).hexdigest()[:8]
            basename += "__" + suffix
        result[path] = output_dir / basename
    return result


def process(inputs: list[Path], *, output: Path, engine: str, crop: str,
            keep_text: bool, require_vector: bool, no_overwrite: bool,
            no_audit: bool, license_path: Path | None, show_excel: bool = False,
            output_format: str = "svg", inkscape_path: Path | None = None) -> dict[str, Any]:
    if output_format not in {"svg", "emf", "both"}:
        raise ExtractionError(f"Unsupported output format: {output_format}")
    formats = ("svg", "emf") if output_format == "both" else (output_format,)
    output.mkdir(parents=True, exist_ok=True)
    backend = backend_for(engine, crop=crop, keep_text=keep_text,
                          license_path=license_path, show_excel=show_excel,
                          inkscape_path=inkscape_path, require_vector=require_vector)
    report: dict[str, Any] = {
        "tool": "excel-charts-to-svg-emf", "version": VERSION,
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "engine": backend.name, "formats": list(formats),
        "workbooks": [], "charts": [], "errors": [], "warnings": [],
    }
    if backend.name == "aspose" and license_path is None:
        report["warnings"].append(
            "Aspose.Cells is running without a license. Evaluation limitations and visual watermarks may apply.")
    assigned = workbook_dirs(inputs, output)
    try:
        for input_file in inputs:
            target_folder = assigned[input_file]
            print(f"Workbook: {input_file}", flush=True)
            observed = 0
            saved = 0
            files_saved = 0
            expected = None
            book = None
            try:
                try:
                    expected = count_ooxml_chart_parts(input_file)
                except (BadZipFile, OSError) as exc:
                    report["warnings"].append(f"Cannot audit OOXML chart count for {input_file}: {exc}")
                book = backend.open(input_file)
                specs = list(backend.enumerate(book))
                observed = len(specs)
                if expected is not None and expected != observed and not no_audit:
                    report["errors"].append(
                        f"Chart completeness mismatch in {input_file}: Excel file has "
                        f"{expected} chart definition(s), backend enumerated {observed}. "
                        "Some charts may be unsupported or charts may be orphaned in the file. "
                        "Use --no-audit to bypass only this structural check.")
                seen: set[str] = set()
                for spec in specs:
                    target_folder.mkdir(parents=True, exist_ok=True)
                    # Allocate the stem exactly once, so .svg and .emf pairs
                    # always use the same collision-safe descriptive name.
                    base = unique_output(target_folder, spec, seen)
                    label = chart_label(spec.title, spec.object_name, spec.position, spec.name_hint)
                    fully_exported = True
                    for fmt in formats:
                        dst = base.with_suffix(f".{fmt}")
                        record = Result(str(input_file), spec.sheet, spec.object_name,
                                        spec.title, spec.chart_type, str(dst), "ok", format=fmt)
                        if no_overwrite and dst.exists():
                            record.status = "error"
                            record.error = f"Output file exists and --no-overwrite is set: {dst}"
                        else:
                            try:
                                # Move only validated files to their final paths.
                                with tempfile.TemporaryDirectory(dir=target_folder,
                                                                 prefix=".render-") as scratch:
                                    intermediate = Path(scratch) / f"chart.{fmt}"
                                    render_method = backend.render(spec, intermediate)
                                    record.render_method = (render_method if isinstance(render_method, str)
                                                            else backend.name)
                                    if fmt == "svg":
                                        info = svg_inspect_and_label(
                                            intermediate, title=label,
                                            description=(f"Workbook: {input_file.name}; "
                                                         f"sheet: {spec.sheet}; chart: {spec.object_name}"),
                                            require_vector=require_vector)
                                    else:
                                        info = emf_inspect(intermediate, require_vector=require_vector)
                                    os.replace(intermediate, dst)
                                record.vector_elements = info.vector_elements
                                record.raster_elements = info.raster_elements
                                files_saved += 1
                                print(f"  OK   {dst.name}", flush=True)
                                if render_method == "duplicate_chart_sheet":
                                    report["warnings"].append(
                                        f"{dst}: used duplicate chart-sheet fallback; chart-sheet layout "
                                        "may differ slightly from the original embedded-chart proportions")
                                if render_method == "inkscape_svg_fallback":
                                    report["warnings"].append(
                                        f"{dst}: EMF made from Excel PDF/SVG via Inkscape "
                                        "because native Excel clipboard EMF was unavailable")
                                if info.raster_elements:
                                    report["warnings"].append(
                                        f"{dst}: includes {info.raster_elements} "
                                        f"raster {'record' if fmt == 'emf' else 'image element'}(s)")
                                if fmt == "emf" and info.emf_plus_comments:
                                    report["warnings"].append(
                                        f"{dst}: contains {info.emf_plus_comments} EMF+ comment(s); "
                                        "their drawing commands are not classified by the validator")
                            except Exception as exc:
                                record.status = "error"
                                record.error = str(exc)
                        report["charts"].append(asdict(record))
                        if record.status == "error":
                            fully_exported = False
                            report["errors"].append(
                                f"{input_file.name} / {spec.sheet} / {label} [{fmt}]: {record.error}")
                            print(f"  FAIL {label} [{fmt}]: {record.error}",
                                  file=sys.stderr, flush=True)
                    if fully_exported:
                        saved += 1
            except Exception as exc:
                report["errors"].append(f"Cannot process {input_file}: {exc}")
                print(f"  FAIL {exc}", file=sys.stderr, flush=True)
            finally:
                if book is not None:
                    try:
                        backend.close_workbook(book)
                    except Exception as exc:
                        report["warnings"].append(f"Could not close {input_file}: {exc}")
                report["workbooks"].append({"source": str(input_file),
                    "output_dir": str(target_folder), "chart_definitions_in_zip": expected,
                    "charts_enumerated": observed, "charts_exported": saved,
                    "files_exported": files_saved})
    finally:
        try:
            backend.close()
        except Exception as exc:
            report["warnings"].append(f"Could not close rendering backend cleanly: {exc}")
    report["summary"] = {"workbooks": len(inputs),
                         "charts_enumerated": sum(x["charts_enumerated"] for x in report["workbooks"]),
                         "charts_exported": sum(x["charts_exported"] for x in report["workbooks"]),
                         "files_exported": sum(x["files_exported"] for x in report["workbooks"]),
                         "errors": len(report["errors"]), "warnings": len(report["warnings"])}
    manifest = output / "manifest.json"
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", dir=output,
                                     prefix=".manifest-", suffix=".json", delete=False) as temp:
        temp_name = temp.name
        json.dump(report, temp, ensure_ascii=False, indent=2)
        temp.write("\n")
    os.replace(temp_name, manifest)
    print(f"\n{report['summary']['files_exported']} file(s) "
          f"({report['summary']['charts_exported']} chart(s) fully exported), "
          f"{report['summary']['errors']} error(s), {report['summary']['warnings']} warning(s).")
    print(f"Manifest: {manifest}")
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Extract Excel charts as validated vector-first SVG, EMF or both.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("inputs", nargs="+", help="Excel files, directories, or wildcards")
    parser.add_argument("-o", "--output", type=Path, default=Path("charts_svg"),
                        help="Output directory; each workbook gets its own subfolder")
    parser.add_argument("-r", "--recursive", action="store_true", help="Scan directory inputs recursively")
    parser.add_argument("--engine", choices=["auto", "excel", "aspose"], default="auto",
                        help="Excel uses native Excel PDF/clipboard; Aspose works on Windows/Linux with a license")
    parser.add_argument("--format", choices=["svg", "emf", "both"], default="svg",
                        help="Output vector format(s); SVG-only preserves previous behavior")
    parser.add_argument("--aspose-license", type=Path, help="Path to an Aspose.Cells license file")
    parser.add_argument("--inkscape", type=Path,
                        help="Optional Inkscape executable used if native Excel EMF clipboard export fails")
    parser.add_argument("--crop", choices=["auto", "none"], default="auto",
                        help="Trim printer whitespace for native Excel PDF export")
    parser.add_argument("--keep-text", action="store_true", help="Keep live SVG text instead of portable font outlines")
    parser.add_argument("--show-excel", action="store_true", help="Show the private automated Excel instance (helps if chart activation fails while hidden)")
    parser.add_argument("--allow-raster", action="store_true",
                        help="Allow raster-only SVG or EMF containers (off by default)")
    parser.add_argument("--no-overwrite", action="store_true", help="Do not replace existing exported files")
    parser.add_argument("--no-audit", action="store_true", help="Do not fail on OOXML chart-count discrepancies")
    parser.add_argument("--version", action="version", version=f"%(prog)s {VERSION}")
    args = parser.parse_args(argv)
    inputs = collect_inputs(args.inputs, args.recursive)
    if not inputs:
        parser.error("No Excel workbook files matched your inputs")
    if args.aspose_license and not args.aspose_license.is_file():
        parser.error(f"Aspose license file not found: {args.aspose_license}")
    if args.inkscape and not args.inkscape.is_file():
        parser.error(f"Inkscape executable not found: {args.inkscape}")
    try:
        report = process(inputs, output=args.output.expanduser().resolve(),
            engine=args.engine, crop=args.crop, keep_text=args.keep_text,
            require_vector=not args.allow_raster, no_overwrite=args.no_overwrite,
            no_audit=args.no_audit, license_path=args.aspose_license,
            show_excel=args.show_excel, output_format=args.format,
            inkscape_path=args.inkscape)
        return 2 if report["errors"] else 0
    except (ExtractionError, OSError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
