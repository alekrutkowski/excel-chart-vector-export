# Contributing

Contributions, tests and bug reports are welcome. Contact: **Alek Rutkowski** (alek.rutkowski@gmail.com).

## Local setup

```bash
python -m venv .venv
python -m pip install -e '.[windows,test]'
python -m pytest -q
python -m compileall -q excel_charts_to_svg.py
```

On Windows, choose the virtual environment's Python (`.venv\Scripts\python`). The Windows-only `pywin32` dependency is automatically skipped outside Windows. On Linux, Inkscape tests are skipped if its binary is not available.

## Testing changes

- Keep public command-line options backward compatible when practical.
- Add regression tests for parsing, file naming, SVG/EMF validation, error paths, and partial failures.
- Do not attach private workbooks. The `examples/` workbook is synthetic and contains 47 embedded Excel charts.
- Excel COM integration requires Windows with desktop Excel installed. Run `scripts/smoke-test-windows.ps1` and report Excel version, Python version, output summary and any sanitized failure messages.
- The native EMF path temporarily uses the Windows clipboard. Avoid testing it while other software is copying to the clipboard.

Before submitting a pull request, run the portable tests and check that README statements and third-party attribution still match the code.

## Release process

1. Update `VERSION` in `excel_charts_to_svg.py`, `version` in `pyproject.toml`, and `CHANGELOG.md`.
2. Run `python -m pytest -q` on Windows and Linux, plus the Excel smoke test on a real Excel installation.
3. Check `python -m pip install -e '.[windows,test]'` and `excel-chart-export --version`.
4. Create a GitHub release tagged `vX.Y.Z`. The release workflow builds the source distribution and wheel and uploads them to that release.
5. Publish results and known limitations in the release notes. Publishing to PyPI is a separate optional step and is not automated here.
