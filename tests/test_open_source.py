"""Portable tests plus optional live integration tests for FOSS rendering engines."""
import json
import os
from pathlib import Path
import shutil
import struct
import sys

import pytest

import excel_charts_to_svg as tool
import libreoffice_uno_worker as worker


FIXTURE = Path(__file__).resolve().parents[1] / "examples" / "Excel_Chart_Extraction_Stress_Test.xlsx"


def test_ooxml_catalog_preserves_titles_and_top_to_bottom_order():
    entries = tool.ooxml_chart_catalog(FIXTURE)
    assert sum(map(len, entries.values())) == 47
    assert len(entries["Sales dashboard"]) == 9
    assert entries["Sales dashboard"][0].title == "Revenue by region"
    assert entries["Sales dashboard"][1].title == "Monthly revenue"
    assert entries["Naming edge cases"][1].title == "Café – naïve Δ, München, 東京"
    assert entries["Science XY"][0].chart_type == "scatterChart"


def test_ooxml_catalog_nonzip_returns_empty(tmp_path):
    p = tmp_path / "old.xls"
    p.write_text("not zip")
    assert tool.ooxml_chart_catalog(p) == {}


def test_resolve_ooxml_relationships():
    assert tool._part_target("xl/worksheets/sheet1.xml", "../drawings/drawing1.xml") == "xl/drawings/drawing1.xml"
    assert tool._part_target("xl/drawings/drawing1.xml", "/xl/charts/chart1.xml") == "xl/charts/chart1.xml"


def test_ooxml_name_mapping_does_not_guess_when_counts_disagree():
    rows = [tool.ChartSpec("Sales dashboard", "Graph 1", "", 1, {})]
    mapped = tool.match_chart_metadata(FIXTURE, rows)
    assert mapped[0].title == ""
    assert mapped[0].name_source == "backend"


def test_ooxml_name_mapping_when_counts_agree():
    items = [tool.ChartSpec("Sales dashboard", "Graph 1", "", i + 1, {}) for i in range(9)]
    mapped = tool.match_chart_metadata(FIXTURE, items)
    assert mapped[0].title == "Revenue by region"
    assert mapped[1].title == "Monthly revenue"
    assert mapped[1].name_source == "ooxml_position"


def test_gnumeric_graph_names_are_generic_without_an_excel_title():
    assert tool.chart_label("", "Graph 3", 3, "") == "Chart 03"
    assert tool.chart_label("", "Graph 3", 3, "Sales") == "Sales"


def create_fake_ssconvert(path: Path, *, charts_per_sheet: dict[str, int]) -> Path:
    # Fake ONLY the command-line subprocess, not the real SVG validator.
    code = '''#!/usr/bin/env python3
import pathlib,sys
assert '--export-graphs' in sys.argv
assert sys.argv[-2].endswith('.xlsx')
template = sys.argv[-1]
for sheet, n in __CHARTS__:
 for i in range(1,n+1):
  output = pathlib.Path(template.replace('%s', sheet).replace('%n', str(i)))
  output.write_text('<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100">'
                    '<path d="M 10 10 L 50 60"/></svg>', encoding='utf-8')
'''.replace("__CHARTS__", repr(list(charts_per_sheet.items())))
    path.write_text(code, encoding="utf-8")
    path.chmod(0o755)
    return path


@pytest.mark.skipif(os.name == "nt", reason="Fake ssconvert uses a Unix shebang")
def test_gnumeric_batch_from_fixture_with_47_real_svg_names(tmp_path):
    mock_cmd = create_fake_ssconvert(tmp_path / "ssconvert", charts_per_sheet={
        "Sales dashboard": 9, "Finance - negatives": 9, "Science XY": 8,
        "Products & categories": 8, "Naming edge cases": 12,
        "Chart only worksheet": 1,
    })
    report = tool.process([FIXTURE], output=tmp_path / "exports", engine="gnumeric",
                          crop="auto", keep_text=False, require_vector=True,
                          no_overwrite=False, no_audit=False, license_path=None,
                          ssconvert_path=mock_cmd)
    assert report["summary"]["charts_exported"] == 47
    assert report["summary"]["files_exported"] == 47
    assert report["summary"]["errors"] == 0
    assert report["charts"][0]["name_source"] == "ooxml_position"
    folder = tmp_path / "exports" / FIXTURE.stem
    assert (folder / "Sales dashboard__Revenue by region.svg").is_file()
    assert (folder / "Naming edge cases__Café – naïve Δ, München, 東京.svg").is_file()
    assert not list(folder.glob("*.emf"))


@pytest.mark.skipif(os.name == "nt", reason="Fake ssconvert uses a Unix shebang")
def test_gnumeric_vector_svg_to_emf_via_real_inkscape(tmp_path):
    if tool.inkscape_executable(None) is None:
        pytest.skip("Optional Inkscape not installed")
    fake = create_fake_ssconvert(tmp_path / "ssconvert", charts_per_sheet={"Sales dashboard": 1})
    backend = tool.GnumericBackend(executable=fake, formats=("svg", "emf"))
    try:
        spec = backend.open(FIXTURE)[0]
        svg = tmp_path / "chart.svg"
        emf = tmp_path / "chart.emf"
        assert backend.render(spec, svg) == "gnumeric_ssconvert_svg"
        assert backend.render(spec, emf) == "gnumeric_inkscape_svg_emf"
        assert tool.svg_inspect_and_label(svg, title="x", description="x", require_vector=True).vector_elements > 0
        assert tool.emf_inspect(emf).vector_elements > 0
    finally:
        backend.close()


@pytest.mark.skipif(os.name == "nt", reason="Fake ssconvert uses a Unix shebang")
def test_gnumeric_emf_without_inkscape_is_clear_error(tmp_path, monkeypatch):
    fake = create_fake_ssconvert(tmp_path / "ssconvert", charts_per_sheet={"Sales dashboard": 1})
    monkeypatch.setattr(tool, "inkscape_executable", lambda *a: None)
    with pytest.raises(tool.ExtractionError, match="EMF requires Inkscape"):
        tool.GnumericBackend(executable=fake, formats=("emf",))


def test_libreoffice_worker_detects_only_chart_ole():
    class Model:
        def supportsService(self, name):
            return name == "com.sun.star.chart2.ChartDocument"
    class Obj:
        ShapeType = "com.sun.star.drawing.OLE2Shape"
    obj = Obj()
    obj.Model = Model()
    assert worker.is_chart(obj)
    obj.ShapeType = "com.sun.star.drawing.GraphicObjectShape"
    assert not worker.is_chart(obj)


def test_libreoffice_worker_uses_original_chart_title():
    class Title:
        String = "Total net revenue"
    class Model:
        def getTitle(self):
            return Title()
    class Shape:
        pass
    shape = Shape()
    shape.Model = Model()
    assert worker.title_from_model(shape) == "Total net revenue"


def test_emf_libreoffice_known_off_by_one_allowed_only_with_opt_in(tmp_path):
    from test_tool import sample_emf
    data = bytearray(sample_emf())
    struct.pack_into("<I", data, 52, 2)
    emf = tmp_path / "lo.emf"
    emf.write_bytes(data)
    with pytest.raises(tool.ExtractionError, match="count"):
        tool.emf_inspect(emf)
    info = tool.emf_inspect(emf, allow_lo_record_discrepancy=True)
    assert info.record_count_difference == 1
    assert info.vector_elements > 0


def test_libreoffice_worker_renders_emf_when_svg_only_fails(tmp_path, monkeypatch):
    calls = []

    def fake_export(smgr, context, shape, target, extension):
        calls.append(extension)
        if extension == "svg":
            raise RuntimeError("GraphicExportFilter refuses SVG")
        target.write_bytes(b"EMF" * 40)

    monkeypatch.setattr(worker, "export_one", fake_export)
    result = worker.export_chart_files(None, None, None, tmp_path, 0, 1, ["svg"])
    assert calls == ["svg", "emf"]
    assert "svg" not in result["files"]
    assert result["errors"]["svg"] == "GraphicExportFilter refuses SVG"
    assert Path(result["fallback_emf"]).is_file()


def test_libreoffice_worker_reuses_requested_emf_for_failed_svg(tmp_path, monkeypatch):
    calls = []

    def fake_export(smgr, context, shape, target, extension):
        calls.append(extension)
        if extension == "svg":
            raise RuntimeError("SVG unavailable")
        target.write_bytes(b"EMF" * 40)

    monkeypatch.setattr(worker, "export_one", fake_export)
    result = worker.export_chart_files(None, None, None, tmp_path, 0, 1, ["svg", "emf"])
    assert calls == ["svg", "emf"]
    assert "fallback_emf" not in result
    assert "emf" in result["files"]


def test_libreoffice_worker_does_not_generate_emf_when_svg_succeeds(tmp_path, monkeypatch):
    calls = []

    def fake_export(smgr, context, shape, target, extension):
        calls.append(extension)
        target.write_text("<svg xmlns='http://www.w3.org/2000/svg'>"
                          "<path d='M 0 0 L 10 10'/></svg>", encoding="utf-8")

    monkeypatch.setattr(worker, "export_one", fake_export)
    result = worker.export_chart_files(None, None, None, tmp_path, 0, 1, ["svg"])
    assert calls == ["svg"]
    assert "fallback_emf" not in result
    assert "svg" in result["files"]


def test_libreoffice_backend_uses_inkscape_for_missing_svg(tmp_path, monkeypatch):
    intermediate = tmp_path / "chart.emf"
    intermediate.write_bytes(b"EMF" * 40)
    backend = object.__new__(tool.LibreOfficeBackend)
    backend.inkscape = Path("/fake/inkscape")
    spec = tool.ChartSpec("Sheet", "Chart", "Title", 1, {
        "files": {"emf": str(intermediate)},
        "errors": {"svg": "LibreOffice failed to render image/svg+xml"},
    })

    def fake_convert(emf, svg, executable):
        assert emf == intermediate
        assert executable == backend.inkscape
        svg.write_text('<svg xmlns="http://www.w3.org/2000/svg" width="100" '
                       'height="100"><path d="M 10 10 L 90 90"/></svg>',
                       encoding="utf-8")

    monkeypatch.setattr(tool, "inkscape_emf_to_svg", fake_convert)
    svg = tmp_path / "out.svg"
    assert backend.render(spec, svg) == "libreoffice_inkscape_emf_svg"
    assert tool.svg_inspect_and_label(svg, title="Title", description="Sheet",
                                      require_vector=True).vector_elements > 0


def test_libreoffice_backend_reports_missing_inkscape_with_native_error(tmp_path):
    intermediate = tmp_path / "chart.emf"
    intermediate.write_bytes(b"EMF" * 40)
    backend = object.__new__(tool.LibreOfficeBackend)
    backend.inkscape = None
    spec = tool.ChartSpec("Sheet", "Chart", "Title", 1, {
        "files": {}, "fallback_emf": str(intermediate),
        "errors": {"svg": "LibreOffice SVG export was unsuccessful"},
    })
    with pytest.raises(tool.ExtractionError, match="requires Inkscape") as exc:
        backend.render(spec, tmp_path / "chart.svg")
    assert "SVG export was unsuccessful" in str(exc.value)


@pytest.mark.skipif(tool.inkscape_executable(None) is None,
                    reason="Real Inkscape not installed")
def test_real_inkscape_emf_to_svg_fallback(tmp_path):
    """Use an actual EMF derived from the bundled xlsx via LibreOffice."""
    if not shutil.which("soffice") and not shutil.which("libreoffice"):
        pytest.skip("Real LibreOffice not installed")
    # Export one of the real chart EMFs through the production backend.
    backend = tool.LibreOfficeBackend(formats=("emf",), timeout=120)
    try:
        chart = next(iter(backend.open(FIXTURE)))
        emf = tmp_path / "chart.emf"
        backend.render(chart, emf)
        info = tool.emf_inspect(emf, allow_lo_record_discrepancy=True)
        assert info.vector_elements > 0
        svg = tmp_path / "chart.svg"
        tool.inkscape_emf_to_svg(emf, svg, tool.inkscape_executable(None))
        svg_info = tool.svg_inspect_and_label(svg, title="real", description="fixture",
                                              require_vector=True)
        assert svg_info.vector_elements > 0
        assert svg_info.external_images == 0
    finally:
        backend.close()


def test_gnumeric_error_if_missing_executable(tmp_path):
    with pytest.raises(tool.ExtractionError, match="ssconvert"):
        tool.GnumericBackend(executable=tmp_path / "no-such-file")


def test_open_source_engine_cli_options_available(capsys):
    with pytest.raises(SystemExit) as ex:
        tool.main(["--help"])
    assert ex.value.code == 0
    text = capsys.readouterr().out
    assert "gnumeric" in text and "libreoffice" in text
    assert "--uno-python" in text and "--ssconvert" in text


@pytest.mark.integration
@pytest.mark.skipif(not os.environ.get("EXCEL_CHART_LIVE_LO"),
                    reason="Set EXCEL_CHART_LIVE_LO=1 to run LibreOffice integration")
def test_live_libreoffice_forced_svg_failure_uses_real_vector_fallback(tmp_path):
    """End-to-end reproduction of the CI failure regardless of LO release.

    Only the test's *temporary copy* of the UNO helper is altered. All real
    LibreOffice/UNO and Inkscape operations run normally for the EMF fallback.
    """
    if tool.inkscape_executable(None) is None:
        pytest.skip("Inkscape is required for forced SVG failure integration test")
    content = Path(worker.__file__).read_text(encoding="utf-8")
    target = ('def export_one(smgr, context, shape, target: Path, extension: str):\n'
              '    import uno\n')
    assert content.count(target) == 1
    altered = content.replace(target, target + (
        '    if extension == "svg":\n'
        '        raise RuntimeError("Simulated missing LibreOffice SVG export filter")\n'), 1)
    temp_helper = tmp_path / "uno_worker_svg_failure.py"
    temp_helper.write_text(altered, encoding="utf-8")
    backend = tool.LibreOfficeBackend(formats=("svg",), timeout=240)
    backend.helper = temp_helper
    try:
        specs = backend.open(FIXTURE)
        assert len(specs) == 47
        first = specs[0]
        assert "svg" not in first.chart["files"]
        assert Path(first.chart["fallback_emf"]).is_file()
        assert "Simulated missing" in first.chart["errors"]["svg"]
        dest = tmp_path / "fallback.svg"
        assert backend.render(first, dest) == "libreoffice_inkscape_emf_svg"
        info = tool.svg_inspect_and_label(dest, title=first.title,
                                          description="forced SVG fallback",
                                          require_vector=True)
        assert info.vector_elements > 0
        assert info.external_images == 0
    finally:
        backend.close()


@pytest.mark.integration
@pytest.mark.skipif(not os.environ.get("EXCEL_CHART_LIVE_LO"),
                    reason="Set EXCEL_CHART_LIVE_LO=1 to run LibreOffice integration")
def test_live_libreoffice_renders_all_47_svg_and_emf(tmp_path):
    report = tool.process([FIXTURE], output=tmp_path / "live", engine="libreoffice",
                          crop="auto", keep_text=False, require_vector=True,
                          no_overwrite=False, no_audit=False, license_path=None,
                          output_format="both", timeout=240)
    assert report["summary"]["charts_enumerated"] == 47
    assert report["summary"]["charts_exported"] == 47
    assert report["summary"]["files_exported"] == 94
    assert report["summary"]["errors"] == 0
    assert all(item["vector_elements"] > 0 for item in report["charts"])
    assert all(item["name_source"] == "libreoffice" for item in report["charts"] if item["title"])
    folder = tmp_path / "live" / FIXTURE.stem
    assert (folder / "Sales dashboard__Revenue by region.svg").is_file()
    assert (folder / "Sales dashboard__Revenue by region.emf").is_file()
