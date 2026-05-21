@echo off
setlocal
cd /d %~dp0
py -m PyInstaller --noconfirm --clean ^
    --name "國土地籍批次查詢工具" ^
    --windowed ^
    --onefile ^
    --paths . ^
    --add-data "config.json;." ^
    --hidden-import openpyxl ^
    --hidden-import pyproj ^
    --exclude-module selenium ^
    --exclude-module bs4 ^
    --exclude-module PySide6 ^
    --exclude-module PyQt5 ^
    --exclude-module PyQt6 ^
    --exclude-module matplotlib ^
    --exclude-module scipy ^
    --exclude-module IPython ^
    --exclude-module notebook ^
    --exclude-module pytest ^
    --exclude-module tkinter.test ^
    land_query.py
echo.
echo === done. exe at dist\國土地籍批次查詢工具.exe ===
endlocal
