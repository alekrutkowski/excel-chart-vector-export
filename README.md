# Excel Chart Vector Export

Export native Excel charts from one or more workbooks to **individual SVG, EMF, or both**, with descriptive filenames and a machine-readable completeness report.

The tool uses **Microsoft Excel on Windows** for the closest match to Excel's original chart appearance. An optional **Aspose.Cells** backend is available on Windows, Linux and macOS. Exports are checked for valid SVG/EMF structure, and raster-only files are rejected by default.

## Requirements

| Backend                          | Where it runs                        | Required software                                                             | Formats     |
| -------------------------------- | ------------------------------------ | ----------------------------------------------------------------------------- | ----------- |
| Native Excel (`--engine excel`)  | Windows, interactive desktop session | Python 3.10+, desktop Microsoft Excel, `pywin32`, `PyMuPDF`                   | SVG and EMF |
| Aspose.Cells (`--engine aspose`) | Windows, Linux, macOS                | Python 3.10+, Java runtime, `aspose-cells`; paid license for unrestricted use | SVG and EMF |

An optional **Inkscape** installation provides an SVG-to-EMF fallback if Excel does not supply an enhanced metafile on the clipboard. The default `--engine auto` tries Excel first on Windows and otherwise uses Aspose.Cells, including a fallback when Excel cannot start. If Aspose runs without a license, evaluation restrictions and watermarks can apply; the manifest warns about this.

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

### Linux or macOS: optional Aspose backend

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
- **Aspose backend:** chart rendering directly to SVG or EMF, using Aspose.Cells.

A genuine vector container can still contain **raster components**, for instance image fills and chart effects. The tool counts recognized vector and raster elements and reports raster content in the manifest. By default, output with no detected vector drawing records is rejected; `--allow-raster` relaxes that check. SVG/EMF validation is **structural**, not a guarantee of visual fidelity or full editability in another application. EMF+ drawing commands are not fully classified, so some vector-rich EMF+ output may be rejected conservatively.

**Important:** Native EMF export **replaces the global Windows clipboard contents**. Avoid copying or pasting in other applications while it runs. This does not apply to SVG-only export or the Aspose backend. If Excel cannot render headlessly, retry with `--show-excel` in a normal interactive session.

## All options

```text
excel-chart-export INPUT [INPUT ...] [-o OUTPUT] [--recursive]
                   [--engine auto|excel|aspose] [--format svg|emf|both]
                   [--aspose-license FILE] [--inkscape FILE]
                   [--crop auto|none] [--keep-text] [--show-excel]
                   [--allow-raster] [--no-overwrite] [--no-audit]
                   [--version]
```

| Option                         | Meaning                                                                  |
| ------------------------------ | ------------------------------------------------------------------------ |
| `--format svg\|emf\|both`      | Output format(s). Default: `svg`, for backward compatibility.            |
| `-o, --output DIR`             | Destination directory. Default: `charts_svg`.                            |
| `-r, --recursive`              | Search within input directories recursively.                             |
| `--engine auto\|excel\|aspose` | Rendering backend. `auto` selects based on platform and availability.    |
| `--aspose-license FILE`        | Aspose.Cells license file for licensed rendering.                        |
| `--inkscape FILE`              | Inkscape executable for the optional EMF fallback.                       |
| `--crop auto\|none`            | Automatic PDF whitespace trimming; `none` preserves the PDF page bounds. |
| `--keep-text`                  | Retain text objects in SVG instead of converting glyphs into paths.      |
| `--show-excel`                 | Make the automated Excel instance visible during export.                 |
| `--allow-raster`               | Accept a container with no detected vector drawing records.              |
| `--no-overwrite`               | Treat an existing target file as an error rather than overwriting it.    |
| `--no-audit`                   | Disable the OOXML chart-definition count cross-check.                    |
| `--version`                    | Print version.                                                           |

Supported inputs: `.xlsx`, `.xlsm`, `.xlsb`, `.xls`, `.xltx`, `.xltm`, `.xlt`. On Windows the native backend searches embedded chart objects, charts in grouped drawing objects when accessible, and standalone chart sheets. Pasted pictures of charts, sparklines and non-chart shapes are outside the scope. The Aspose backend enumerates charts in worksheets; standalone chart sheets and complex grouped items may not be available there.

## Validation and failure reporting

The tool writes `manifest.json` to the output root. It includes the source workbook, chart and worksheet names, output filename, target format, rendering method, record/element counts, failure messages, and totals.

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

On Linux, `pywin32` is skipped by the platform dependency marker. The Inkscape tests skip when Inkscape is unavailable. GitHub Actions runs the portable tests on Windows and Ubuntu. **GitHub-hosted Windows runners do not supply Microsoft Excel**, so Excel automation is validated separately using the [Windows smoke test](scripts/smoke-test-windows.ps1).

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
| Font or layout differs from Excel                          | On Windows prefer `--engine excel`; install required fonts. Aspose or Inkscape conversion may differ visually.                                                   |
| SVG rejected as raster-only                                | Check if the original chart uses embedded pictures or effects. Use `--allow-raster` only if raster-based output is acceptable.                                   |
| A chart is missing                                         | Review `manifest.json`, try `--show-excel`, and check unsupported chart structures. `--no-audit` suppresses the count check but does not recover omitted charts. |
| Some files exported, others failed                         | Examine the per-chart `errors` and `render_method` fields in `manifest.json`. The tool keeps successful files.                                                   |

## Privacy, limits and security

No web service is used to render or upload the workbook. The program invokes local software (Excel, PyMuPDF, optional Inkscape or Aspose). Be cautious with untrusted Excel files; macro execution is disabled in the native backend, but Office files and third-party renderers still need security patching. Do not commit workbooks containing private business data, or an Aspose license file, to GitHub.

Excel on Windows runs in a **separate automated application instance**. The tool reads specified files rather than enumerating all files currently open in Excel. This project exports *rendered chart graphics*, not the underlying editable Excel chart object or series data.

## License and dependency acknowledgments

**Project source code: MIT License**, copyright 2026 **Alek Rutkowski** ([contact](mailto:alek.rutkowski@gmail.com)). See [LICENSE](LICENSE). The MIT license covers this repository's source code only. It does **not** grant licenses to Microsoft Excel, PyMuPDF, Aspose.Cells, Inkscape, or any other separate software; users and redistributors must follow those terms.

The project acknowledges these software projects and their maintainers:

| Component                                     | Role                                                   | License or terms                                                                           | Project                                              |
| --------------------------------------------- | ------------------------------------------------------ | ------------------------------------------------------------------------------------------ | ---------------------------------------------------- |
| **PyMuPDF / MuPDF** (Artifex Software)        | Excel PDF to SVG; PDF inspection                       | **AGPL-3.0 or commercial license**. Important for closed-source/commercial redistribution. | https://pymupdf.io/licensing                         |
| **pywin32** (Mark Hammond and contributors)   | Windows Excel COM automation and EMF clipboard access  | Python Software Foundation License                                                         | https://pypi.org/project/pywin32/                    |
| **Aspose.Cells for Python via Java** (Aspose) | Optional Excel-independent SVG/EMF backend             | Proprietary commercial licensing; evaluation limitations                                   | https://docs.aspose.com/cells/python-java/licensing/ |
| **JPype1** (Steve Menard and contributors)    | Java bridge used by the optional Aspose backend        | Apache-2.0                                                                                 | https://pypi.org/project/jpype1/                     |
| **Inkscape** (Inkscape contributors)          | Optional SVG-to-EMF fallback                           | Open-source GPL terms (refer to upstream distribution)                                     | https://inkscape.org/                                |
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

## Development and publication

See [CONTRIBUTING.md](CONTRIBUTING.md), [CHANGELOG.md](CHANGELOG.md), [SECURITY.md](SECURITY.md), and the [publication checklist](PUBLISHING.md). The repository includes a Python package entry point, editable installation, a portable test suite, GitHub Actions checks and a tagged-release packaging workflow. **Publishing to GitHub is deliberately left to the repository owner**; no remote has been created or modified.
