@echo off
title Terraria Paintings Scraper
cd /d "%~dp0"

echo ============================================
echo   Terraria Wiki Paintings Scraper
echo ============================================
echo.

:: Install / update required packages silently
echo Checking Python packages ...
pip install -q requests beautifulsoup4 openpyxl Pillow gspread google-auth 2>nul
echo.

echo Scraping paintings from terraria.wiki.gg ...
echo This may take a minute or two on the first run.
echo.

python scrape_paintings.py

echo.
if %ERRORLEVEL% EQU 0 (
    echo ============================================
    echo   Done! Opening paintings.xlsx ...
    echo ============================================
    echo.
    echo File: %~dp0paintings.xlsx
    echo.
    start "" "%~dp0paintings.xlsx"
) else (
    echo ============================================
    echo   Something went wrong. See error above.
    echo ============================================
)

pause
