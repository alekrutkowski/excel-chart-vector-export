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
