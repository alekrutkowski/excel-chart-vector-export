# Publish this repository on GitHub

The provided ZIP is a complete working tree. It deliberately contains no `.git` directory and is not connected to a GitHub repository.

## Before making it public

1. Check that `README.md` and `LICENSE` match your intended terms. The project code is MIT-licensed, while PyMuPDF is AGPL/commercial and Aspose.Cells is proprietary; these third-party terms are separate.
2. Review all files for proprietary or personally identifying data. The bundled example workbook is synthetic. Never commit Aspose license files.
3. Run portable tests:

   ```powershell
   py -3 -m pip install -e ".[windows,test]"
   py -3 -m pytest -q
   ```

4. On Windows with Microsoft Excel, run the native end-to-end smoke test:

   ```powershell
   .\scripts\smoke-test-windows.ps1
   ```

   Check `manifest.json`, visually compare a sample of SVG and EMF outputs against the original charts, and record any chart types that render differently. This manual integration check is important because GitHub-hosted test runners do not provide Microsoft Excel.

## Create the GitHub repository

With [Git](https://git-scm.com/) and the optional [GitHub CLI](https://cli.github.com/) installed, open PowerShell inside the extracted directory:

```powershell
git init
git add .
git commit -m "Prepare Excel Chart Vector Export v1.2.1 for GitHub"
gh auth login
gh repo create excel-chart-vector-export --public --source . --remote origin --push
```

If you already have a GitHub repository, configure its remote and push instead. Run `git status` before pushing to verify which files are tracked. The `.gitignore` excludes common caches, generated outputs and `*.lic` license files.

## Releases

1. Keep the version synchronized in `pyproject.toml` and `excel_charts_to_svg.py`; update `CHANGELOG.md`.
2. Push all changes and verify the **Tests** GitHub Actions workflow passes.
3. Create a GitHub release with a matching tag, e.g. `v1.2.1`. The release workflow builds a source distribution and wheel and attaches them to that release.
4. Describe any outstanding live Excel limitations in the public release notes. This project currently remains beta until full Windows Excel smoke testing is confirmed.

The workflow does not publish to PyPI. PyPI publication is a separate decision and should be taken only after checking package availability, security and all applicable licensing obligations.
