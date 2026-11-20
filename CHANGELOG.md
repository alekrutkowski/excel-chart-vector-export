# Changelog

## [1.3.0] - 2026-11-19

- Add fully open-source `--engine libreoffice`: headless LibreOffice UNO chart-by-chart SVG and native EMF export, using an isolated profile and the real `GraphicExportFilter`.
- Add `--engine gnumeric`: `ssconvert` vector SVG batch export with optional Inkscape conversion to vector EMF.
- Prefer LibreOffice/Gnumeric before commercial Aspose when auto-selecting a backend without Windows Excel.
- Add OOXML worksheet/chart relationship metadata to improve naming for renderers that do not return chart titles; record naming provenance.
- Accept only LibreOffice's reproducible one-record EMF header discrepancy, with an explicit warning; keep strict checks elsewhere.
- Add genuine LibreOffice integration coverage (47 charts, 94 vector files, zero errors on Linux test environment).
- Extend documentation and GitHub CI with open-source install, usage, license and limitation details.
- Keep the original Microsoft Excel and Aspose rendering modes available.


## [1.2.1] - 2026-11-19

- Prepare a structured GitHub repository with MIT license, author metadata, Python packaging, CI, release workflow, and a synthetic chart test workbook.
- Clarify third-party licenses, especially PyMuPDF (AGPL/commercial) and Aspose.Cells (commercial).
- Improve naming of untitled charts whose Excel object name is generic (`Chart` or `Chart N`).
- Reject external resource links and executable content in exported SVG files.
- Add regression tests for SVG safety and the included 47-chart workbook.
- Keep the existing SVG/EMF CLI behavior and output structure.

## [1.2.0] - 2026-11-19

- Introduce `--format svg|emf|both`.
- Add Excel clipboard-based native EMF export, structural EMF validation, and optional Inkscape fallback.
- Save paired SVG and EMF outputs using identical descriptive stems and report format-specific failures.

## [1.1.0] - 2026-11-19

- Address the worksheet-instead-of-chart PDF export problem through chart activation and a temporary duplicated chart-sheet fallback.
- Add regression tests for the multi-page worksheet issue.

## [1.0.0] - 2026-11-19

- First batch-export implementation using native Excel and optional Aspose.Cells rendering.
- Add descriptive chart naming, validation, OOXML chart-count audit and a JSON manifest.

> Historical version descriptions summarize implementation changes. An automated or mocked test pass is not a guarantee of successful live Excel COM export.
