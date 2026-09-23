# PDF → Excel

Minimalistische Windows-Desktop-App zum Einlesen von PDF-Rechnungen und Export nach Excel.

## Funktionen
- PDFs per Drag & Drop
- mehrere PDFs gleichzeitig
- automatische Erkennung von Rechnungsnummer, Datum, Fälligkeit, Soll, Haben und Saldo
- 30 Tage Zahlungsziel als Fallback
- dunkles Design
- Tabelle direkt bearbeitbar
- Excel-Export
- GitHub Actions baut automatisch eine Windows-EXE

## Start
```bash
pip install -r requirements.txt
python app.py
```

## EXE
Nach einem Push auf `main`: GitHub → Actions → Windows EXE → Artefakt herunterladen.
