@echo off
pyinstaller --onefile --hidden-import openpyxl --hidden-import pywin32 --hidden-import PyQt5 TaxHandler.py TaxHandlerComponents.py
pause
