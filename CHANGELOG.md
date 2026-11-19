# Changelog

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
