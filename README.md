# Excel Chart Vector Export

Export native Excel charts from one or more workbooks to **individual SVG, EMF, or both**, with descriptive filenames and a machine-readable completeness report.

**No Microsoft Excel or paid libraries are required** when using the new **LibreOffice Calc** backend (SVG and EMF) or **Gnumeric** backend (SVG, with optional Inkscape conversion to EMF). The original Microsoft Excel backend remains available on Windows for the closest match to Excel's chart appearance, and Aspose.Cells remains optional. Exports undergo SVG/EMF structural validation; raster-only files are rejected by default.

**Verified on LibreOffice 25.2 (Linux):** the bundled 47-chart test workbook produced **47 vector SVGs and 47 vector EMFs**, with zero extraction errors. A later GitHub Actions run on Ubuntu 26.04 initially failed to produce SVGs because the workflow installed Calc without the Draw SVG filter package. Version 1.3.1 installs that dependency in CI and can fall back to converting LibreOffice's vector EMF with Inkscape. Live GitHub verification of v1.3.1 is still required. This does not guarantee the same result with other Excel files or LibreOffice versions. Gnumeric integration is unit-tested but has not been run against an installed `ssconvert` in this test environment.

## Requirements

| Backend | Where it runs | Required software | Formats |
| --- | --- | --- | --- |
| **LibreOffice (`--engine libreoffice`)** | Linux, Windows; macOS with compatible UNO Python | Python 3.10+, LibreOffice Calc **and Draw** (SVG filter), `soffice`, Python UNO bridge; optional Inkscape fallback | **SVG and EMF**, directly; SVG may use EMF conversion fallback |
| **Gnumeric (`--engine gnumeric`)** | Linux, including WSL2 on Windows | Python 3.10+, Gnumeric's `ssconvert`; optional Inkscape | **SVG**; EMF via Inkscape |
| Native Excel (`--engine excel`) | Windows, interactive desktop session | Python 3.10+, desktop Microsoft Excel, `pywin32`, `PyMuPDF` | SVG and EMF |
| Aspose.Cells (`--engine aspose`) | Windows, Linux, macOS | Python 3.10+, Java runtime, `aspose-cells`; paid license for unrestricted use | SVG and EMF |

**Inkscape** provides SVG-to-EMF conversion for Gnumeric and native Excel, plus optional EMF-to-SVG fallback when LibreOffice's native SVG filter is missing or fails. The default `--engine auto` selects the first backend that can be initialized: **Excel on Windows**, then **LibreOffice**, **Gnumeric**, and **Aspose.Cells**. If the chosen renderer later fails to import or draw a chart, the failure is reported; `auto` does not mix different renderers inside a workbook. Explicitly choose `--engine` when reproducibility matters. Aspose evaluation restrictions and watermarks can apply without a commercial license.

## Quick start: LibreOffice (fully open source)

On Ubuntu / WSL2 (native Linux Python and LibreOffice):

```bash
sudo apt update
sudo apt install libreoffice-calc libreoffice-draw python3-uno
python3 -m venv .venv
.venv/bin/python -m pip install -e .
.venv/bin/python excel_charts_to_svg.py ./Report.xlsx \
    --engine libreoffice --format both -o ./ExportedCharts
```

The tool launches an **isolated headless LibreOffice instance**, opens the supplied workbook read-only with macro execution and link updating disabled, and exports each chart shape through LibreOffice's `GraphicExportFilter`. **LibreOffice Draw is needed for the native SVG export filter** on Debian-family systems: installing Calc alone may allow EMF export but fail for every SVG. SVG and EMF normally use native vector export, not screenshots or file-extension tricks. If native SVG fails while EMF succeeds and Inkscape is installed, the tool automatically converts the vector EMF into a real SVG (with structural validation). The manifest records `render_method: libreoffice_inkscape_emf_svg` and a warning when that fallback is used. Install Inkscape with `sudo apt install inkscape`, or supply `--inkscape PATH`. The fallback can change text appearance or other drawing details, so inspect important charts. The LibreOffice UNO bridge often belongs to the *system* Python rather than your venv; the script finds a compatible interpreter automatically (usually `/usr/bin/python3` on Ubuntu). If necessary, specify `--uno-python /path/to/python3`, `--soffice /path/to/soffice`, and adjust `--timeout` (default: 300 seconds per workbook).

On Windows, install LibreOffice and use the corresponding `soffice.exe` with `--soffice`. You must also provide a Python interpreter that can `import uno`, using `--uno-python` if it is not detected automatically; Windows LibreOffice installers do not necessarily include a matching, standalone Python interpreter. The normal Windows Python used to run the exporter can still be in a virtual environment. See [LibreOffice's UNO developer documentation](https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1GraphicExportFilter.html).

### Gnumeric (simplest headless SVG export)

On Ubuntu / WSL2:

```bash
sudo apt install gnumeric
python3 excel_charts_to_svg.py ./Report.xlsx \
    --engine gnumeric --format svg -o ./ExportedCharts
```

To request **both SVG and EMF**, also install Inkscape and use `--format both`. Gnumeric's `ssconvert --export-graphs` produces the SVGs; Inkscape converts each genuine SVG to a genuine EMF. Without Inkscape, `--engine gnumeric --format emf|both` stops with a clear dependency error. Gnumeric cannot export EMF natively through `ssconvert`.

```bash
sudo apt install gnumeric inkscape
python3 excel_charts_to_svg.py ./Report.xlsx \
    --engine gnumeric --format both -o ./ExportedCharts
```

The Gnumeric backend names charts by matching Gnumeric's sheet/index output with chart titles recorded in OOXML, **only if each sheet has the expected number of charts**. This assumes the renderer's top-to-bottom, left-to-right ordering is unchanged. When the mapping cannot be trusted, generic numbered names are used rather than silently assigning the wrong title. The `name_source` field in `manifest.json` reveals the naming method. Gnumeric may not import every chart type supported by Excel; count mismatches are reported as errors.

## Quick start: Windows

Download the repository ZIP, extract it, and open PowerShell in the repository directory. With Python installed:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python -m pip install -e ".[windows]"
```

Export all charts from **one workbook** to SVG:

```powershell
.\.venv\Scripts\python .\excel_charts_to_svg.py "C:\Data\Report.xlsx" -o "C:\Data\ExportedCharts"
```

Export **both** vector formats:

```powershell
.\.venv\Scripts\python .\excel_charts_to_svg.py "C:\Data\Report.xlsx" --format both -o "C:\Data\ExportedCharts"
```

Or run it through the installed console command, after activating the virtual environment:

```powershell
.\.venv\Scripts\Activate.ps1
excel-chart-export "C:\Data\Report.xlsx" --format emf -o "C:\Data\ExportedCharts"
```

If your PowerShell execution policy prevents activating the virtual environment, use the direct `\.venv\Scripts\python` commands above. You do not have to activate the environment.

The legacy dependency installation also remains supported: `py -3 -m pip install -r requirements-windows.txt`. You can run the standalone `excel_charts_to_svg.py` file without installing the package itself, provided its backend dependencies are installed.

### Multiple workbooks

```powershell
.\.venv\Scripts\python .\excel_charts_to_svg.py "C:\Reports" --recursive --format both -o "C:\Charts"
```

Multiple explicit input paths and globs are also supported. The tool uses the files **you pass**; it does not automatically scan every workbook currently open in Excel. The native backend starts a separate Excel automation instance and opens each input read-only.

### Optional Aspose backend

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -e '.[aspose]'
.venv/bin/python excel_charts_to_svg.py ./Report.xlsx \
    --engine aspose --format both \
    --aspose-license /path/to/Aspose.Cells.lic -o ./ExportedCharts
```

Aspose is proprietary and its evaluation mode can add watermarks or impose other restrictions. See [Aspose licensing](https://docs.aspose.com/cells/python-java/licensing/). The software needs a suitable Java runtime. As of November 2026, the current `aspose-cells` release on PyPI declares support only through Python 3.13; use a supported interpreter for this optional backend. Output from Aspose may differ visually from desktop Excel, especially in fonts, effects and unsupported charts.

## Output

For `Report.xlsx` with two charts on the `Sales` worksheet:

```text
ExportedCharts/
  Report/
    Sales__Revenue by region.svg
    Sales__Revenue by region.emf
    Sales__Net margin.svg
    Sales__Net margin.emf
  manifest.json
```

Files are named from their **worksheet and chart title**. Untitled charts use a meaningful object name or data-series names when available, then a numbered fallback. Windows-invalid characters are replaced; Unicode is retained. Names that collide, including case-only collisions, receive suffixes such as `__02`. Workbooks with identical stems are placed in separate, deterministic hashed subfolders.

The original workbooks are opened read-only, macro execution is disabled for native Excel automation, external link updates are suppressed, and **no changes are saved**. Temporary chart sheets used for fallback rendering exist only in Excel memory, and the source file is closed without saving.

### What "vector" means here

- **SVG (Excel backend):** Excel chart exported as a one-page PDF, converted into SVG vector drawing elements by PyMuPDF. By default, text becomes paths for portability; `--keep-text` preserves live SVG text at the expense of font portability.
- **EMF (Excel backend):** Excel's `CopyPicture` requests a metafile, and the binary Windows `CF_ENHMETAFILE` clipboard format is captured and validated. On failure, the optional Inkscape fallback converts a vector SVG to EMF.
- **LibreOffice backend:** reads `.xlsx` with Calc and renders individual chart shapes directly to **SVG** (`image/svg+xml`) and **EMF** (`image/x-emf`) with UNO's `GraphicExportFilter`. On Ubuntu/Debian install `libreoffice-draw` for the native SVG filter. When SVG fails but EMF succeeds, installed Inkscape can convert the real vector EMF to a valid SVG. The vector paths and text reflect LibreOffice's rendering, which can differ from Excel; fallback conversion adds another potential source of differences.
- **Gnumeric backend:** exports individual SVGs using `ssconvert --export-graphs`. For EMF, Inkscape converts these SVGs to vector EMFs; Gnumeric does not offer a native EMF chart exporter.
- **Aspose backend:** chart rendering directly to SVG or EMF, using Aspose.Cells.

A genuine vector container can still contain **raster components**, for instance image fills and chart effects. The tool counts recognized vector and raster elements and reports raster content in the manifest. By default, output with no detected vector drawing records is rejected; `--allow-raster` relaxes that check. SVG/EMF validation is **structural**, not a guarantee of visual fidelity or full editability in another application. EMF+ drawing commands are not fully classified, so some vector-rich EMF+ output may be rejected conservatively. Some LibreOffice-generated EMFs declare **one fewer record in the header** than is present in their complete, EOF-terminated record stream; the LibreOffice backend accepts only this precise discrepancy and reports a warning. All other backends use strict EMF record-count checking.

**Important:** Native EMF export **replaces the global Windows clipboard contents**. Avoid copying or pasting in other applications while it runs. This does not apply to SVG-only export or the Aspose backend. If Excel cannot render headlessly, retry with `--show-excel` in a normal interactive session.

## All options

```text
excel-chart-export INPUT [INPUT ...] [-o OUTPUT] [--recursive]
                   [--engine auto|excel|libreoffice|gnumeric|aspose]
                   [--format svg|emf|both] [--aspose-license FILE]
                   [--inkscape FILE] [--ssconvert FILE] [--soffice FILE]
                   [--uno-python FILE] [--timeout SECONDS]
                   [--crop auto|none] [--keep-text] [--show-excel]
                   [--allow-raster] [--no-overwrite] [--no-audit]
                   [--version]
```

| Option                         | Meaning                                                                  |
| ------------------------------ | ------------------------------------------------------------------------ |
| `--format svg\|emf\|both`      | Output format(s). Default: `svg`, for backward compatibility.            |
| `-o, --output DIR`             | Destination directory. Default: `charts_svg`.                            |
| `-r, --recursive`              | Search within input directories recursively.                             |
| `--engine auto\|excel\|libreoffice\|gnumeric\|aspose` | Rendering backend. `auto` tries installed/usable engines in that order, with Excel first on Windows. |
| `--aspose-license FILE`        | Aspose.Cells license file for licensed rendering.                        |
| `--inkscape FILE` | Inkscape executable for Gnumeric/Excel EMF export and LibreOffice EMF-to-SVG fallback. |
| `--ssconvert FILE` | Gnumeric `ssconvert` executable (default: search `PATH`). |
| `--soffice FILE` | LibreOffice `soffice` executable (default: search `PATH`). |
| `--uno-python FILE` | Python interpreter with LibreOffice's UNO bridge available. |
| `--timeout SECONDS` | Per-workbook subprocess timeout for LibreOffice/Gnumeric (default: 300). |
| `--crop auto\|none`            | Automatic PDF whitespace trimming; `none` preserves the PDF page bounds. |
| `--keep-text`                  | Retain text objects in SVG instead of converting glyphs into paths.      |
| `--show-excel`                 | Make the automated Excel instance visible during export.                 |
| `--allow-raster`               | Accept a container with no detected vector drawing records.              |
| `--no-overwrite`               | Treat an existing target file as an error rather than overwriting it.    |
| `--no-audit`                   | Disable the OOXML chart-definition count cross-check.                    |
| `--version`                    | Print version.                                                           |

Accepted extensions: `.xlsx`, `.xlsm`, `.xlsb`, `.xls`, `.xltx`, `.xltm`, `.xlt`. **Actual import support depends on the chosen renderer.** The native Excel backend searches embedded charts, chart-bearing grouped drawing objects where accessible, and standalone chart sheets. LibreOffice enumerates embedded Calc chart shapes (including some grouped charts); Gnumeric exports imported graph objects. Neither open-source backend currently promises extraction from standalone Excel chart sheets. Pasted images of charts, sparklines, and non-chart shapes are out of scope. The Aspose backend enumerates worksheet charts; chart sheets and complex groups may be unsupported.

## Validation and failure reporting

The tool writes `manifest.json` to the output root. It includes the source workbook, chart and worksheet names, filename, target format, rendering method, SVG/EMF vector and raster element counts, **naming source** (`backend`, `libreoffice`, `ooxml_position`, or `gnumeric_sheet_index`), errors, warnings, and totals.

```json
{
  "summary": {
    "workbooks": 1,
    "charts_enumerated": 47,
    "charts_exported": 47,
    "files_exported": 94,
    "errors": 0,
    "warnings": 0
  }
}
```

This is an **illustrative successful outcome** for the supplied 47-chart test workbook with `--format both`; it is not a claimed result from a live Windows Excel test.

`charts_exported` counts charts for which all requested formats succeeded. `files_exported` counts actual saved output files. When extracting both formats, a failed EMF does not delete a successfully exported SVG. Files are saved atomically after validation. The program exits with code **0** if there are no errors, or **2** if processing/export fails.

For OOXML workbooks, an independent cross-check counts chart references and compares them with charts discovered by the backend. If the counts differ, the mismatch is reported as an error, but available charts are still processed. `--no-audit` disables only the comparison. The audit is a useful coverage check, not proof that every chart is visually correct.

## Test workbook and testing

[`examples/Excel_Chart_Extraction_Stress_Test.xlsx`](examples/Excel_Chart_Extraction_Stress_Test.xlsx) is a synthetic workbook containing **47 embedded chart objects**, across eight chart families, including unusual Unicode titles, duplicate names, negative values and varied dimensions. It does not include hidden worksheets or true standalone Excel chart sheets.

Run portable automated tests:

```bash
python -m pip install -e '.[windows,test]'
python -m pytest -q
```

For the **real LibreOffice end-to-end test** on Linux:

```bash
sudo apt install libreoffice-calc libreoffice-draw python3-uno
EXCEL_CHART_LIVE_LO=1 python -m pytest -q -m integration
```

A separate GitHub Actions job runs this test on Ubuntu. It checks the synthetic workbook's **47 chart objects and 94 vector files**. Gnumeric's subprocess and naming logic have automated tests, but the real `ssconvert` rendering results depend on the installed version; run `scripts/smoke-test-gnumeric.sh` to assess its import coverage on your machine. The script **fails loudly** if Gnumeric misses any of the fixture's 47 charts.

On Linux, `pywin32` is skipped by the platform dependency marker. The Inkscape tests skip when Inkscape is unavailable. GitHub Actions runs portable tests on Windows and Ubuntu and live LibreOffice integration on Ubuntu 24.04 and 26.04, with Calc, Draw, UNO and Inkscape installed. The integration suite forces a native SVG failure for a separate fallback regression test. **GitHub-hosted Windows runners do not supply Microsoft Excel**, so Excel automation is validated separately using the [Windows smoke test](scripts/smoke-test-windows.ps1).

On a Windows machine with Excel installed:

```powershell
.\scripts\smoke-test-windows.ps1
```

The smoke test runs both formats on the included workbook and checks that 47 charts and 94 files were exported with no errors. A failure is useful regression evidence. It may also reveal an Excel-version-specific incompatibility, since live COM/clipboard handling is not emulated in CI.

## Troubleshooting

| Symptom                                                    | What to try                                                                                                                                                      |
| ---------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Expected a single-chart PDF, but Excel printed N page(s)` | Upgrade from v1.0 to this release; use `--show-excel`. The activated-chart method and temporary chart-sheet fallback are intended to prevent worksheet printing. |
| No `CF_ENHMETAFILE` clipboard data                         | Run in an interactive Windows desktop session, do not copy to the clipboard during export, try `--show-excel`, or install Inkscape for fallback conversion.      |
| Font or layout differs from Excel | For closest original styling on Windows prefer `--engine excel`; install fonts. LibreOffice, Gnumeric and Aspose use their own layout engines and may differ. |
| SVG rejected as raster-only                                | Check if the original chart uses embedded pictures or effects. Use `--allow-raster` only if raster-based output is acceptable.                                   |
| A chart is missing                                         | Review `manifest.json`, try `--show-excel`, and check unsupported chart structures. `--no-audit` suppresses the count check but does not recover omitted charts. |
| Some files exported, others failed | Examine per-chart `errors`, `render_method`, and `name_source` in `manifest.json`. Successful outputs are retained. |
| LibreOffice: UNO bridge not found | Install `python3-uno` (Linux), or supply a compatible interpreter using `--uno-python`. `soffice` must also be installed. |
| LibreOffice: all SVG exports fail but EMF works | On Debian/Ubuntu, install `libreoffice-draw` as well as `libreoffice-calc`. Inkscape (`sudo apt install inkscape`) is an optional real-vector fallback if the native SVG exporter still fails. Check the `render_method` in `manifest.json`. |
| Gnumeric: `ssconvert` missing | Install `gnumeric` on Linux/WSL2, or supply the executable via `--ssconvert`. |
| EMF unavailable with Gnumeric | Install Inkscape or choose `--format svg`. Gnumeric itself does not export EMF. |
| LibreOffice EMF record-count warning | A known one-record mismatch was observed in valid, vector-rich EMF output; it is accepted only for this backend. |

## Privacy, limits and security

No web service is used to render or upload the workbook. The program invokes local software (Excel, LibreOffice, Gnumeric, PyMuPDF, optional Inkscape, or Aspose). LibreOffice uses an isolated, loopback-only UNO listener and a temporary user profile. Be cautious with untrusted Excel files; macro execution is disabled in the native backend, but Office files and third-party renderers still need security patching. Do not commit workbooks containing private business data, or an Aspose license file, to GitHub.

Excel on Windows runs in a **separate automated application instance**. The tool reads specified files rather than enumerating all files currently open in Excel. This project exports *rendered chart graphics*, not the underlying editable Excel chart object or series data.

## License and dependency acknowledgments

**Project source code: MIT License**, copyright 2026 **Alek Rutkowski** ([contact](mailto:alek.rutkowski@gmail.com)). See [LICENSE](LICENSE). The MIT license covers this repository's source code only. It does **not** grant licenses to Microsoft Excel, PyMuPDF, Aspose.Cells, Inkscape, or any other separate software; users and redistributors must follow those terms.

The project acknowledges these software projects and their maintainers:

| Component                                     | Role                                                   | License or terms                                                                           | Project                                              |
| --------------------------------------------- | ------------------------------------------------------ | ------------------------------------------------------------------------------------------ | ---------------------------------------------------- |
| **PyMuPDF / MuPDF** (Artifex Software)        | Excel PDF to SVG; PDF inspection                       | **AGPL-3.0 or commercial license**. Important for closed-source/commercial redistribution. | https://pymupdf.io/licensing                         |
| **pywin32** (Mark Hammond and contributors)   | Windows Excel COM automation and EMF clipboard access  | Python Software Foundation License                                                         | https://pypi.org/project/pywin32/                    |
| **LibreOffice Calc / Draw / UNO** (The Document Foundation and contributors) | Open-source Excel-independent SVG/EMF renderer | MPL 2.0 and LGPL 3.0+ as distributed by the project | https://www.libreoffice.org/about-us/licenses/ |
| **Gnumeric / ssconvert** (Gnumeric and GNOME contributors) | Open-source XLSX reader and SVG chart exporter | GPL v2/v3 (see distribution's licensing) | https://www.gnumeric.org/ |
| **Aspose.Cells for Python via Java** (Aspose) | Optional Excel-independent SVG/EMF backend             | Proprietary commercial licensing; evaluation limitations                                   | https://docs.aspose.com/cells/python-java/licensing/ |
| **JPype1** (Steve Menard and contributors)    | Java bridge used by the optional Aspose backend        | Apache-2.0                                                                                 | https://pypi.org/project/jpype1/                     |
| **Inkscape** (Inkscape contributors)          | Optional SVG-to-EMF and EMF-to-SVG fallback                           | Open-source GPL terms (refer to upstream distribution)                                     | https://inkscape.org/                                |
| **pytest** (pytest contributors)              | Automated tests; development-only                      | MIT                                                                                        | https://docs.pytest.org/en/stable/license.html       |
| **setuptools** (Python Packaging Authority)   | Python build/install tooling; build-time only          | MIT                                                                                        | https://pypi.org/project/setuptools/                 |
| **Microsoft Excel** (Microsoft)               | Native chart renderer on Windows; installed separately | Commercial Microsoft software terms                                                        | https://www.microsoft.com/microsoft-365/excel        |

This repository does not vendor, redistribute or modify third-party binaries. Installing optional dependencies can introduce further transitive dependencies with their own licenses. Verify the terms for your intended use. This licensing summary is informational and is not legal advice.

### Relevant API documentation

- [Excel Chart.ExportAsFixedFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.exportasfixedformat)
- [Excel Chart.CopyPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.copypicture)
- [Excel ChartObject.CopyPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobject.copypicture)
- [pywin32 GetClipboardData](https://mhammond.github.io/pywin32/win32clipboard__GetClipboardData_meth.html)
- [Aspose.Cells Chart.toImage](https://reference.aspose.com/cells/python-java/asposecells.api/chart/)
- [PyMuPDF Page.get_svg_image](https://pymupdf.readthedocs.io/en/latest/page.html)
- [Inkscape command-line use](https://wiki.inkscape.org/wiki/Using_the_Command_Line)
- [Gnumeric `ssconvert` manual (including graph SVG exports)](https://manpages.debian.org/testing/gnumeric/ssconvert.1.en.html)
- [LibreOffice `GraphicExportFilter` UNO API](https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1GraphicExportFilter.html)
- [LibreOffice Calc Guide: chart SVG/EMF export](https://books.libreoffice.org/en/CG262/CG26206-CreatingChartsAndGraphs.html)

## Development and publication

See [CONTRIBUTING.md](CONTRIBUTING.md), [CHANGELOG.md](CHANGELOG.md), [SECURITY.md](SECURITY.md), and the [publication checklist](PUBLISHING.md). The repository includes a Python package entry point, editable installation, a portable test suite, GitHub Actions checks and a tagged-release packaging workflow. **Publishing to GitHub is deliberately left to the repository owner**; no remote has been created or modified.
