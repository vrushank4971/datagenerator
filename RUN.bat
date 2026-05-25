@echo off
REM Double-click this file to generate the 3 PDFs.
REM Requires Python to be installed and added to PATH.

echo.
echo ========================================
echo   Chapter 10 PDF Generator
echo ========================================
echo.

python split_responses.py

echo.
echo Press any key to close...
pause >nul
