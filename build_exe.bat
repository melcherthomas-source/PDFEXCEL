@echo off
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --clean --onefile --windowed --name PDF_zu_Excel app.py
echo Fertig: dist\PDF_zu_Excel.exe
pause
