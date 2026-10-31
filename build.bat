@echo off
chcp 65001 >nul
pip install -r requirements-dev.txt
pyinstaller --onefile --windowed --name Shareholders shareholders.py
echo.
echo Done. Output: dist\Shareholders.exe
pause
