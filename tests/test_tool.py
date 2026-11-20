"""Tests for the portable portions and fake-backend end-to-end pipeline.

Run from project root with ``python -m pytest -q``.
"""
import json
from pathlib import Path
import sys
from zipfile import ZipFile

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import excel_charts_to_svg as tool


def test_readable_and_safe_filenames():
    assert tool.clean_filename(' Annual: Revenue / Growth? ') == 'Annual- Revenue - Growth'
    assert tool.clean_filename('CON') == '_CON'
    assert tool.clean_filename(' Évolution – dépenses ') == 'Évolution – dépenses'
    assert tool.clean_filename('X' * 100, 10) == 'X' * 10
    assert tool.clean_filename('') == 'Untitled'


def test_input_resolution_and_recursive(tmp_path):
    a = tmp_path / 'a.xlsx'
    a.touch()
    (tmp_path / '~$a.xlsx').touch()
    sub = tmp_path / 'nested'
    sub.mkdir()
    b = sub / 'b.xlsb'
    b.touch()
    assert tool.collect_inputs([str(tmp_path)], False) == [a]
    assert set(tool.collect_inputs([str(tmp_path), str(a)], True)) == {a, b}
    assert tool.collect_inputs([str(tmp_path / '*.xlsx')], False) == [a]


def test_ooxml_count(tmp_path):
    p = tmp_path / 'book.xlsx'
    with ZipFile(p, 'w') as z:
        z.writestr('xl/charts/chart1.xml', 'some text')
        z.writestr('xl/charts/chartEx2.xml', 'some text')
        z.writestr('xl/charts/chartStyle1.xml', 'not a chart')
        z.writestr('xl/drawings/chart1.xml', 'not a chart')
    assert tool.count_ooxml_chart_parts(p) == 2


def test_dedupes_names_and_workbook_dirs(tmp_path):
    folder = tmp_path / 'output'
    spec = tool.ChartSpec('P&L', 'Chart 1', 'Sales: 2026', 1, None)
    seen = set()
    assert tool.unique_output(folder, spec, seen).name == 'P&L__Sales- 2026.svg'
    assert tool.unique_output(folder, spec, seen).name == 'P&L__Sales- 2026__02.svg'
    a = tmp_path / 'a' / 'Report.xlsx'
    b = tmp_path / 'b' / 'Report.xlsx'
    dirs = tool.workbook_dirs([a, b], folder)
    assert dirs[a] != dirs[b]


def test_pdf_export_is_vector_and_cropped(tmp_path):
    fitz = pytest.importorskip('pymupdf')
    pdf = tmp_path / 'one-page.pdf'
    svg = tmp_path / 'chart.svg'
    document = fitz.open()
    page = document.new_page(width=612, height=792)
    shape = page.new_shape()
    shape.draw_rect(fitz.Rect(100, 120, 410, 310))
    shape.finish(color=(0.1, 0.2, 0.3), fill=(0.95, 0.95, 0.95))
    shape.commit()
    page.insert_text((128, 170), 'Test Revenue 2026', fontsize=19)
    document.save(pdf)
    document.close()
    tool.pdf_to_svg(pdf, svg, crop='auto', keep_text=False)
    info = tool.svg_inspect_and_label(svg, title='Test Revenue 2026',
                                      description='Unit test example', require_vector=True)
    assert info.vector_elements > 0
    assert info.raster_elements == 0
    # This tests actual bounds, not just a syntactically-valid SVG.
    assert float(info.width.removesuffix('pt')) < 420
    assert float(info.height.removesuffix('pt')) < 290
    assert 'Test Revenue 2026' in svg.read_text(encoding='utf-8')


def test_raster_only_rejected(tmp_path):
    svg = tmp_path / 'only_raster.svg'
    svg.write_text('<svg xmlns="http://www.w3.org/2000/svg">'
                   '<image href="data:image/png;base64,iVBORw0KGgo=" width="1" height="1"/>'
                   '</svg>', encoding='utf-8')
    with pytest.raises(tool.ExtractionError, match='raster'):
        tool.svg_inspect_and_label(svg, title='x', description='x', require_vector=True)
    assert tool.svg_inspect_and_label(svg, title='x', description='x',
                                      require_vector=False).raster_elements == 1


def test_external_image_rejected(tmp_path):
    svg = tmp_path / 'external.svg'
    svg.write_text('<svg xmlns="http://www.w3.org/2000/svg">'
                   '<rect width="10" height="10"/><image href="relative.png"/>'
                   '</svg>', encoding='utf-8')
    with pytest.raises(tool.ExtractionError, match='external image'):
        tool.svg_inspect_and_label(svg, title='x', description='x', require_vector=True)


def test_manifest_and_atomic_processing(tmp_path, monkeypatch):
    source = tmp_path / 'source.xlsx'
    with ZipFile(source, 'w') as z:
        z.writestr('xl/charts/chart1.xml', '<c:chart/>')
        z.writestr('xl/charts/chart2.xml', '<c:chart/>')

    class FakeBackend:
        name = 'fake'
        def open(self, path):
            return object()
        def enumerate(self, book):
            return iter([
                tool.ChartSpec('Sales', 'Chart 1', 'Net sales', 1, None),
                tool.ChartSpec('Sales', 'Chart 2', 'Net sales', 2, None),
            ])
        def render(self, spec, dst):
            dst.write_text('<svg xmlns="http://www.w3.org/2000/svg" width="300" '
                           'height="200"><path d="M 1,1 L 5,5"/></svg>', encoding='utf-8')
        def close_workbook(self, book):
            pass
        def close(self):
            pass

    monkeypatch.setattr(tool, 'backend_for', lambda *a, **k: FakeBackend())
    dest = tmp_path / 'out'
    report = tool.process([source], output=dest, engine='auto', crop='auto',
            keep_text=False, require_vector=True, no_overwrite=False,
            no_audit=False, license_path=None)
    assert report['summary']['charts_exported'] == 2
    assert report['summary']['errors'] == 0
    assert (dest / 'source' / 'Sales__Net sales.svg').is_file()
    assert (dest / 'source' / 'Sales__Net sales__02.svg').is_file()
    assert json.loads((dest / 'manifest.json').read_text())['summary']['charts_exported'] == 2


def test_chart_completeness_fails_loudly(tmp_path, monkeypatch):
    source = tmp_path / 'source.xlsx'
    with ZipFile(source, 'w') as z:
        z.writestr('xl/charts/chart1.xml', '')
    class FakeBackend:
        name = 'fake'
        def open(self, path): return object()
        def enumerate(self, book): return iter([])
        def close_workbook(self, book): pass
        def close(self): pass
    monkeypatch.setattr(tool, 'backend_for', lambda *a, **k: FakeBackend())
    report = tool.process([source], output=tmp_path / 'out', engine='auto',
            crop='auto', keep_text=False, require_vector=True,
            no_overwrite=False, no_audit=False, license_path=None)
    assert len(report['errors']) == 1
    assert 'mismatch' in report['errors'][0]


def test_ooxml_relationships_exclude_orphaned_parts(tmp_path):
    p = tmp_path / 'book.xlsx'
    rels = '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="../charts/chartEx2.xml"/>
    </Relationships>'''
    with ZipFile(p, 'w') as z:
        z.writestr('xl/charts/chart1.xml', '')
        z.writestr('xl/charts/chartEx2.xml', '')
        z.writestr('xl/charts/chart3.xml', '')
        z.writestr('xl/drawings/_rels/drawing1.xml.rels', rels)
    assert tool.count_ooxml_chart_parts(p) == 2


def make_sample_pdf(path, pages=1):
    pymupdf = pytest.importorskip('pymupdf')
    with pymupdf.open() as document:
        for i in range(pages):
            page = document.new_page(width=210, height=145)
            page.draw_rect(pymupdf.Rect(15, 20, 190, 115), color=(0, 0, 0))
            page.insert_text((25, 60), "Chart is here", fontsize=14)
        document.save(path)


def test_hidden_chart_is_unhidden_only_temporarily(tmp_path):
    from types import SimpleNamespace as NS
    app = NS(ActiveChart=None)
    class ParentSheet:
        Visible = 0
        def Activate(self):
            assert self.Visible == -1

    class ChartObject:
        Visible = False
        def Activate(self):
            assert self.Visible is True
            app.ActiveChart = chart

    class FakeChart:
        def ExportAsFixedFormat(self, typ, filename, *args):
            if sheet.Visible != -1 or object_.Visible is False:
                raise RuntimeError('Cannot print a hidden chart')
            make_sample_pdf(filename)

    sheet = ParentSheet()
    object_ = ChartObject()
    chart = FakeChart()
    spec = tool.ChartSpec('Hidden', 'Chart 1', 'Hidden chart', 1,
                          chart, host_sheet=sheet, chart_object=object_)
    engine = object.__new__(tool.ExcelBackend)
    engine.crop, engine.keep_text, engine.app = 'auto', False, app
    target = tmp_path / 'chart.svg'
    assert engine.render(spec, target) == 'activated_chart'
    assert target.exists()
    assert sheet.Visible == 0
    assert object_.Visible is False
    info = tool.svg_inspect_and_label(target, title='Hidden chart',
                                      description='Test', require_vector=True)
    assert info.vector_elements > 0


def test_regression_activation_avoids_multi_page_worksheet_pdf(tmp_path):
    """Regression test for 47/47 failures in real-world Excel test workbook."""
    from types import SimpleNamespace as NS
    app = NS(ActiveChart=None)
    class Chart:
        def ExportAsFixedFormat(self, typ, filename, *args):
            # Without explicitly activating this embedded chart, Excel might
            # silently export the entire worksheet (11-19 PDF pages).
            make_sample_pdf(filename, pages=1 if app.ActiveChart is self else 11)

    chart = Chart()
    class Obj:
        Visible = True
        def Activate(self):
            app.ActiveChart = chart
        def Duplicate(self):
            raise AssertionError('fallback is unnecessary when activation works')

    class Sheet:
        Visible = -1
        def Activate(self):
            pass

    spec = tool.ChartSpec('Dashboard', 'Chart 1', 'Revenue by region', 1,
                          chart, host_sheet=Sheet(), chart_object=Obj())
    engine = object.__new__(tool.ExcelBackend)
    engine.crop, engine.keep_text, engine.app = 'auto', False, app
    svg = tmp_path / 'chart.svg'
    assert engine.render(spec, svg) == 'activated_chart'
    assert svg.exists()


def test_fallback_to_duplicate_chart_sheet_after_multi_page_pdf(tmp_path):
    from types import SimpleNamespace as NS
    app = NS(ActiveChart=None)
    calls = []
    class BadChart:
        def ExportAsFixedFormat(self, typ, filename, *args):
            calls.append('bad-pdf')
            make_sample_pdf(filename, pages=19)

    class OneChartSheet:
        def Activate(self):
            calls.append('chart-sheet-activate')
        def ExportAsFixedFormat(self, typ, filename, *args):
            calls.append('good-pdf')
            make_sample_pdf(filename, pages=1)
        def Delete(self):
            calls.append('chart-sheet-delete')

    original = BadChart()
    new_sheet = OneChartSheet()
    class DuplicatedChart:
        def Location(self, where, name):
            assert where == tool.XL_LOCATION_AS_NEW_SHEET
            assert len(name) < 32
            calls.append('location')
            return new_sheet

    class Duplicate:
        Chart = DuplicatedChart()
        def Delete(self):
            raise AssertionError('No duplicate chart object remains to delete after moving')

    class Obj:
        Visible = True
        def Activate(self):
            app.ActiveChart = original
        def Duplicate(self):
            calls.append('duplicate')
            return Duplicate()

    class Sheet:
        Visible = -1
        def Activate(self):
            pass

    spec = tool.ChartSpec('Dashboard', 'Chart 1', 'Revenue by region', 1,
                          original, host_sheet=Sheet(), chart_object=Obj())
    engine = object.__new__(tool.ExcelBackend)
    engine.crop, engine.keep_text, engine.app = 'auto', False, app
    svg = tmp_path / 'chart.svg'
    assert engine.render(spec, svg) == 'duplicate_chart_sheet'
    assert svg.is_file()
    assert calls == ['bad-pdf', 'duplicate', 'location', 'chart-sheet-activate',
                     'good-pdf', 'chart-sheet-delete']


def test_both_pdf_methods_fail_loudly(tmp_path):
    from types import SimpleNamespace as NS
    app = NS(ActiveChart=None)
    class Chart:
        def ExportAsFixedFormat(self, typ, filename, *args):
            make_sample_pdf(filename, pages=11)

    chart = Chart()
    class Obj:
        Visible = True
        def Activate(self):
            app.ActiveChart = chart
        def Duplicate(self):
            raise RuntimeError('worksheet is protected')
    class Sheet:
        Visible = -1
        def Activate(self):
            pass
    spec = tool.ChartSpec('Sheet', 'Chart 1', 'Title', 1, chart,
                          host_sheet=Sheet(), chart_object=Obj())
    engine = object.__new__(tool.ExcelBackend)
    engine.crop, engine.keep_text, engine.app = 'auto', False, app
    dest = tmp_path / 'chart.svg'
    with pytest.raises(tool.ExtractionError, match='Both chart-only PDF methods failed.*11.*protected'):
        engine.render(spec, dest)
    assert not dest.exists()


def test_standalone_chart_sheet_is_activated_before_export(tmp_path):
    from types import SimpleNamespace as NS
    app = NS(ActiveChart=None)
    class Chart:
        Visible = -1
        def Activate(self):
            app.ActiveChart = self
        def ExportAsFixedFormat(self, typ, filename, *args):
            assert app.ActiveChart is self
            make_sample_pdf(filename)
    chart = Chart()
    spec = tool.ChartSpec('ChartSheet', 'ChartSheet', 'Overview', 1,
                          chart, host_sheet=chart)
    engine = object.__new__(tool.ExcelBackend)
    engine.crop, engine.keep_text, engine.app = 'auto', False, app
    assert engine.render(spec, tmp_path / 'chart.svg') == 'activated_chart'


def test_chart_shapelayer_and_group_scan_without_duplicates():
    from types import SimpleNamespace as NS

    def make_chart_shape(name):
        return NS(Name=name, Type=3, HasChart=-1,
                  Chart=NS(HasTitle=False, ChartType=4))

    shape_a = make_chart_shape('Chart 1')
    shape_b = make_chart_shape('Grouped Chart')
    items = NS(Count=1, Item=lambda index: shape_b)
    group = NS(Type=6, GroupItems=items)
    chart_objects = type('ChartObjects', (), {
        'Count': 1,
        '__call__': lambda self, index=None: NS(
            Name='Chart 1', Chart=NS(HasTitle=False, ChartType=4))})()
    ws = NS(Name='Sheet1', Shapes=[shape_a, group],
            ChartObjects=lambda: chart_objects)
    wb = NS(Worksheets=[ws], Charts=[])
    specs = list(tool.excel_specs(wb))
    assert [x.object_name for x in specs] == ['Chart 1', 'Grouped Chart']


def test_untitled_charts_use_series_names():
    assert tool.chart_label('', 'Chart 12', 12, 'Revenue and costs') == 'Revenue and costs'
    assert tool.chart_label('', 'Profit analysis', 12, 'Revenue') == 'Profit analysis'
    assert tool.descriptive_series_names(['Series 1', 'Revenue', 'Cost']) == 'Revenue and Cost'
    assert tool.descriptive_series_names(['=Sheet1!$B$1', 'Revenue']) == 'Revenue'


def sample_emf(*, raster_only=False):
    """A minimal valid EMF with a rectangle or raster record and a real EOF."""
    import struct
    header = bytearray(88)
    struct.pack_into('<II', header, 0, 1, 88)
    header[40:44] = b' EMF'
    struct.pack_into('<I', header, 44, 0x10000)
    if raster_only:
        # EMR_STRETCHDIBITS is raster-only; dummy fixed-width payload suffices
        # for the parser to classify the record type without interpreting it.
        drawn = struct.pack('<II', 81, 24) + b'\x00' * 16
    else:
        drawn = struct.pack('<IIiiii', 43, 24, 0, 0, 400, 200)
    eof = struct.pack('<IIIII', 14, 20, 0, 0, 20)
    struct.pack_into('<II', header, 48, len(header) + len(drawn) + len(eof), 3)
    return bytes(header) + drawn + eof


def test_emf_parser_recognizes_real_binary_vector_records(tmp_path):
    path = tmp_path / 'sample.emf'
    path.write_bytes(sample_emf())
    info = tool.emf_inspect(path)
    assert (info.records, info.vector_elements, info.raster_elements) == (3, 1, 0)
    assert path.read_bytes()[40:44] == b' EMF'


def test_emf_parser_rejects_fake_png_and_corrupt_records(tmp_path):
    path = tmp_path / 'fake.emf'
    path.write_bytes(b'\x89PNG\r\n\x1a\n' + bytes(200))
    with pytest.raises(tool.ExtractionError, match='header'):
        tool.emf_inspect(path)
    path.write_bytes(sample_emf()[:-8])
    with pytest.raises(tool.ExtractionError, match='length'):
        tool.emf_inspect(path)
    fake = bytearray(sample_emf())
    fake[40:44] = b'PNG!'
    path.write_bytes(fake)
    with pytest.raises(tool.ExtractionError, match='signature'):
        tool.emf_inspect(path)


def test_emf_raster_only_rejected_by_default(tmp_path):
    path = tmp_path / 'raster.emf'
    path.write_bytes(sample_emf(raster_only=True))
    with pytest.raises(tool.ExtractionError, match='No classic GDI vector'):
        tool.emf_inspect(path)
    assert tool.emf_inspect(path, require_vector=False).raster_elements == 1


def test_emf_parser_rejects_record_count_mismatch(tmp_path):
    import struct
    raw = bytearray(sample_emf())
    struct.pack_into('<I', raw, 52, 4)
    path = tmp_path / 'bad.emf'
    path.write_bytes(raw)
    with pytest.raises(tool.ExtractionError, match='record count'):
        tool.emf_inspect(path)


def test_native_excel_emf_method_uses_real_metafile_clipboard(tmp_path, monkeypatch):
    from types import SimpleNamespace as NS
    app = NS(ActiveChart=None)
    calls = []
    chart = NS()
    class ChartObj:
        Visible = True
        def Activate(self):
            app.ActiveChart = chart
        def CopyPicture(self, **kwargs):
            calls.append(kwargs)
    obj = ChartObj()
    class Sheet:
        Visible = -1
        def Activate(self): pass
    spec = tool.ChartSpec('Dashboard', 'Chart 1', 'Revenue', 1, chart,
                          host_sheet=Sheet(), chart_object=obj)
    engine = object.__new__(tool.ExcelBackend)
    engine.app, engine.inkscape_path, engine.require_vector = app, None, True
    monkeypatch.setattr(tool, 'clipboard_sequence', lambda: None)
    monkeypatch.setattr(tool, 'get_emf_clipboard_bytes', lambda: sample_emf())
    path = tmp_path / 'chart.emf'
    assert engine.render(spec, path) == 'excel_clipboard_emf'
    assert tool.emf_inspect(path).vector_elements == 1
    assert calls == [dict(Appearance=tool.XL_SCREEN, Format=tool.XL_PICTURE)]


def test_native_emf_clipboard_retries_and_closes_when_busy(monkeypatch):
    from types import ModuleType
    fake = ModuleType('win32clipboard')
    calls = []
    def opened():
        calls.append('open')
        if calls.count('open') == 1:
            raise RuntimeError('clipboard is busy')
    def close():
        calls.append('close')
    fake.OpenClipboard = opened
    fake.CloseClipboard = close
    fake.IsClipboardFormatAvailable = lambda x: x == 14
    fake.GetClipboardData = lambda x: sample_emf()
    monkeypatch.setitem(sys.modules, 'win32clipboard', fake)
    monkeypatch.setattr(tool.time, 'sleep', lambda s: None)
    assert tool.get_emf_clipboard_bytes() == sample_emf()
    assert calls == ['open', 'open', 'close']


def test_both_formats_share_names_and_report_files(tmp_path, monkeypatch):
    source = tmp_path / 'Reports.xlsx'
    with ZipFile(source, 'w') as z:
        z.writestr('xl/charts/chart1.xml', '')
        z.writestr('xl/charts/chart2.xml', '')

    class FakeBackend:
        name = 'fake'
        def open(self, path): return object()
        def enumerate(self, book):
            return iter([tool.ChartSpec('Dashboard', 'Chart 1', 'Revenue', 1, None),
                         tool.ChartSpec('Dashboard', 'Chart 2', 'Revenue', 2, None)])
        def render(self, spec, dst):
            if dst.suffix == '.emf':
                dst.write_bytes(sample_emf())
                return 'excel_clipboard_emf'
            dst.write_text('<svg xmlns="http://www.w3.org/2000/svg" width="250" '
                           'height="150"><path d="M 0,0 L 10,10"/></svg>')
            return 'activated_chart'
        def close_workbook(self, book): pass
        def close(self): pass
    monkeypatch.setattr(tool, 'backend_for', lambda *a, **k: FakeBackend())
    report = tool.process([source], output=tmp_path / 'out', engine='excel',
                          crop='auto', keep_text=False, require_vector=True,
                          no_overwrite=False, no_audit=False, license_path=None,
                          output_format='both')
    folder = tmp_path / 'out' / 'Reports'
    assert sorted(p.name for p in folder.iterdir()) == [
        'Dashboard__Revenue.emf', 'Dashboard__Revenue.svg',
        'Dashboard__Revenue__02.emf', 'Dashboard__Revenue__02.svg']
    assert report['summary']['charts_exported'] == 2
    assert report['summary']['files_exported'] == 4
    assert report['summary']['errors'] == 0
    assert sorted({row['format'] for row in report['charts']}) == ['emf', 'svg']
    assert all(row['vector_elements'] > 0 for row in report['charts'])


def test_one_format_failure_does_not_lose_other_output(tmp_path, monkeypatch):
    source = tmp_path / 'Report.xlsx'
    with ZipFile(source, 'w') as z:
        z.writestr('xl/charts/chart1.xml', '')
    class FakeBackend:
        name = 'fake'
        def open(self, path): return object()
        def enumerate(self, book):
            return iter([tool.ChartSpec('Dashboard', 'Chart 1', 'Revenue', 1, None)])
        def render(self, spec, dst):
            if dst.suffix == '.emf':
                dst.write_bytes(b'\x89PNG' + b'0' * 200)
            else:
                dst.write_text('<svg xmlns="http://www.w3.org/2000/svg">'
                               '<path d="M 0,0 L 10,10"/></svg>')
        def close_workbook(self, book): pass
        def close(self): pass
    monkeypatch.setattr(tool, 'backend_for', lambda *a, **k: FakeBackend())
    report = tool.process([source], output=tmp_path / 'out', engine='excel',
                          crop='auto', keep_text=False, require_vector=True,
                          no_overwrite=False, no_audit=False, license_path=None,
                          output_format='both')
    folder = tmp_path / 'out' / 'Report'
    assert (folder / 'Dashboard__Revenue.svg').exists()
    assert not (folder / 'Dashboard__Revenue.emf').exists()
    assert report['summary']['charts_exported'] == 0
    assert report['summary']['files_exported'] == 1
    assert report['summary']['errors'] == 1
    assert '[emf]' in report['errors'][0]


def test_svg_mode_still_default(tmp_path, monkeypatch):
    # Existing signature callers need not provide output_format.
    spec = tool.ChartSpec('A', 'Chart 1', 'KPI', 1, None)
    assert tool.unique_output(tmp_path, spec, set()).suffix == '.svg'
    assert tool.VERSION == '1.3.1'


def test_real_inkscape_converts_svg_to_genuine_vector_emf(tmp_path):
    inkscape = tool.inkscape_executable(None)
    if inkscape is None:
        pytest.skip('Inkscape is not installed')
    svg = tmp_path / 'chart.svg'
    emf = tmp_path / 'chart.emf'
    svg.write_text('<svg xmlns="http://www.w3.org/2000/svg" width="320" height="240">'
                   '<rect width="300" height="220" fill="#eeeeee"/>'
                   '<path d="M 10 200 L 75 120 L 180 80" stroke="red" '
                   'fill="none" stroke-width="4"/></svg>')
    tool.inkscape_svg_to_emf(svg, emf, inkscape)
    info = tool.emf_inspect(emf)
    assert info.vector_elements > 0
    assert info.raster_elements == 0


def test_excel_emf_fallback_uses_inkscape_and_records_method(tmp_path, monkeypatch):
    inkscape = tool.inkscape_executable(None)
    if inkscape is None:
        pytest.skip('Inkscape is not installed')
    engine = object.__new__(tool.ExcelBackend)
    engine.inkscape_path, engine.require_vector = inkscape, True
    def fail_native(spec, dst):
        raise tool.ExtractionError('No CF_ENHMETAFILE available')
    def svg_export(spec, dst):
        dst.write_text('<svg xmlns="http://www.w3.org/2000/svg" width="100" '
                       'height="100"><rect width="80" height="60"/></svg>')
        return 'activated_chart'
    monkeypatch.setattr(engine, '_render_native_emf', fail_native)
    monkeypatch.setattr(engine, '_render_svg', svg_export)
    spec = tool.ChartSpec('S', 'Chart 1', 'Revenue', 1, None)
    out = tmp_path / 'chart.emf'
    assert engine.render(spec, out) == 'inkscape_svg_fallback'
    assert tool.emf_inspect(out).vector_elements > 0


def test_title_fallback_uses_series_not_generic_chart_object():
    assert tool.chart_label('', 'Chart', 4, 'Revenue') == 'Revenue'
    assert tool.chart_label('', 'Chart 1', 4, 'Revenue') == 'Revenue'
    assert tool.chart_label('', 'Chart', 4, '') == 'Chart 04'
    assert tool.chart_label('', 'Custom chart', 4, 'Revenue') == 'Custom chart'


@pytest.mark.parametrize('payload, warning', [
    ('<use href="https://example.org/a.svg#icon"/>', 'external image'),
    ('<rect style="fill:url(https://example.org/a.svg)"/>', 'external image'),
    ('<script>notAllowed()</script>', 'Unsafe SVG element'),
    ('<foreignObject/>', 'Unsafe SVG element'),
    ('<rect onload="notAllowed()"/>', 'event handler'),
])
def test_svg_rejects_unsafe_external_or_executable_content(tmp_path, payload, warning):
    chart = tmp_path / 'bad.svg'
    chart.write_text('<svg xmlns="http://www.w3.org/2000/svg">'
                     '<path d="M 0 0 L 1 1"/>' + payload + '</svg>')
    with pytest.raises(tool.ExtractionError, match=warning):
        tool.svg_inspect_and_label(chart, title='test', description='test', require_vector=True)


def test_svg_internal_references_are_supported(tmp_path):
    chart = tmp_path / 'good.svg'
    chart.write_text('<svg xmlns="http://www.w3.org/2000/svg">'
                     '<defs><path id="g" d="M 0 0 L 1 1"/></defs>'
                     '<use href="#g"/><rect fill="url(#gradient)" width="10" height="10"/>'
                     '</svg>')
    assert tool.svg_inspect_and_label(chart, title='test', description='test',
                                      require_vector=True).external_images == 0


def test_packaged_stress_workbook_has_47_excel_charts():
    fixture = Path(__file__).resolve().parents[1] / 'examples' / 'Excel_Chart_Extraction_Stress_Test.xlsx'
    assert fixture.is_file()
    assert tool.count_ooxml_chart_parts(fixture) == 47
