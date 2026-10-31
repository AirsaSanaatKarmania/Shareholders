@echo off
chcp 65001 >nul
pip install -r requirements.txt -r requirements-dev.txt
pyinstaller --onefile --windowed --name Shareholders --hidden-import openpyxl shareholders.py
echo.
echo Done. Output: dist\Shareholders.exe
pause
