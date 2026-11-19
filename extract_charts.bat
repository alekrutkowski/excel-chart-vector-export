@echo off
setlocal
py -3 "%~dp0excel_charts_to_svg.py" %*
if errorlevel 1 (
    echo.
    echo Extraction failed. See the errors above and charts_svg\manifest.json.
    pause
)
