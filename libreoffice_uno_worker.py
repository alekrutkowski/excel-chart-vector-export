#!/usr/bin/env python3
"""Helper run with a Python interpreter that supports LibreOffice's `uno` module.

Executed by excel_charts_to_svg.py in an isolated LibreOffice headless instance.
This module depends only on the Python standard library and the UNO bridge.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
import time
import traceback


def property_value(name, value):
    from com.sun.star.beans import PropertyValue
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item


def connect(port: int, *, timeout: float = 30):
    import uno
    context = uno.getComponentContext()
    resolver = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", context)
    last_error = None
    for _ in range(max(1, int(timeout / 0.2))):
        try:
            remote = resolver.resolve(
                f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
            return remote, remote.ServiceManager
        except Exception as exc:
            last_error = exc
            time.sleep(0.2)
    raise RuntimeError(f"Cannot connect to isolated LibreOffice session: {last_error}")


def is_chart(shape):
    try:
        if shape.ShapeType != "com.sun.star.drawing.OLE2Shape":
            return False
        model = shape.Model
        return bool(model and model.supportsService("com.sun.star.chart2.ChartDocument"))
    except Exception:
        # Only verified chart models are exported: ordinary OLE drawings are excluded.
        return False


def iter_charts(page):
    for i in range(page.getCount()):
        shape = page.getByIndex(i)
        if is_chart(shape):
            yield shape
        else:
            # A group may contain charts; some LibreOffice versions expose
            # grouped drawing shapes as an XShapes collection.
            try:
                if shape.ShapeType == "com.sun.star.drawing.GroupShape":
                    yield from iter_charts(shape)
            except Exception:
                pass


def title_from_model(shape):
    """LibreOffice's old chart model exposes a plain-text title when available."""
    try:
        model = shape.Model
        title = model.getTitle()
        if title is None:
            return ""
        # The legacy XChartDocument title is a drawing text shape with String.
        try:
            value = str(title.String or "").strip()
            if value:
                return value
        except Exception:
            pass
        # Chart2 Title.getText() is an XTitle interface on some releases;
        # property 'Text' or 'getText' may be available on others.
        for attr in ("getText", "Text"):
            try:
                prop = getattr(title, attr)
                segments = prop() if callable(prop) else prop
                pieces = [str(part.String) for part in segments]
                if pieces:
                    return "".join(pieces).strip()
            except Exception:
                continue
    except Exception:
        pass
    return ""


def export_one(smgr, context, shape, target: Path, extension: str):
    import uno
    media = {"svg": "image/svg+xml", "emf": "image/x-emf"}[extension]
    exporter = smgr.createInstanceWithContext(
        "com.sun.star.drawing.GraphicExportFilter", context)
    if not exporter.supportsMimeType(media):
        raise RuntimeError(f"LibreOffice GraphicExportFilter does not support {media}")
    exporter.setSourceDocument(shape)
    target.parent.mkdir(parents=True, exist_ok=True)
    options = (property_value("URL", uno.systemPathToFileUrl(str(target))),
               property_value("MediaType", media))
    if not exporter.filter(options) or not target.is_file() or target.stat().st_size < 100:
        raise RuntimeError(f"LibreOffice failed to render {media}")


def run(port: int, input_file: Path, output: Path, formats: list[str]) -> dict:
    import uno
    context, smgr = connect(port)
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", context)
    # MacroExecutionMode.NEVER_EXECUTE = 4. UpdateDocMode.NO_UPDATE = 0.
    props = (property_value("Hidden", True),
             property_value("ReadOnly", True),
             property_value("MacroExecutionMode", 4),
             property_value("UpdateDocMode", 0),
             property_value("AsTemplate", False),
             property_value("AllowLinkUpdate", False))
    document = desktop.loadComponentFromURL(
        uno.systemPathToFileUrl(str(input_file)), "_blank", 0, props)
    if document is None:
        raise RuntimeError(f"LibreOffice could not open {input_file}")
    result = {"charts": []}
    try:
        sheets = document.getSheets()
        for s in range(sheets.getCount()):
            sheet = sheets.getByIndex(s)
            sheet_name = str(sheet.getName())
            chart_index = 0
            for shape in iter_charts(sheet.getDrawPage()):
                chart_index += 1
                title = title_from_model(shape)
                record = {"sheet": sheet_name, "object_name": str(shape.Name or "Chart"),
                          "title": title, "position": chart_index,
                          "chart_type": "", "files": {}, "errors": {}}
                for extension in formats:
                    outfile = output / f"chart_{s:04d}_{chart_index:05d}.{extension}"
                    try:
                        export_one(smgr, context, shape, outfile, extension)
                        record["files"][extension] = str(outfile)
                    except Exception as exc:
                        record["errors"][extension] = str(exc)
                result["charts"].append(record)
    finally:
        # Close the read-only document, never save or alter the original file.
        document.close(True)
    return result


def shutdown_session(port: int) -> None:
    """Shut down the isolated LibreOffice instance, including on Windows.

    Unlike terminating soffice.exe alone, this also works when the launcher
    has handed control to a separate soffice.bin process.
    """
    context, smgr = connect(port, timeout=5)
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", context)
    desktop.terminate()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, required=True)
    parser.add_argument("--input", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--formats", choices=("svg", "emf", "both"))
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--shutdown", action="store_true")
    args = parser.parse_args()
    if args.shutdown:
        try:
            shutdown_session(args.port)
            return 0
        except Exception:
            traceback.print_exc(file=sys.stderr)
            return 2
    if not all((args.input, args.output, args.formats, args.manifest)):
        parser.error("--input, --output, --formats and --manifest are required for exports")
    args.output.mkdir(parents=True, exist_ok=True)
    formats = ["svg", "emf"] if args.formats == "both" else [args.formats]
    try:
        result = run(args.port, args.input.resolve(), args.output.resolve(), formats)
        args.manifest.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        traceback.print_exc(file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
