
import pdfplumber
import csv
import os
import glob
import sys

print("Programm gestartet.")
# Print script location and current working directory for clarity
script_dir = os.path.dirname(os.path.abspath(__file__))
cwd = os.getcwd()
print(f"[INFO] Script location: {script_dir}")
print(f"[INFO] Current working directory: {cwd}")
try:
    # Dynamically find the first PDF in the script's directory, compatible with PyInstaller/frozen executables
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = script_dir
    print(f"[INFO] PDF-Suchverzeichnis (base_path): {base_path}")
    print("[INFO] Suche nach PDF-Dateien in:", os.path.join(base_path, '*.pdf'))
    pdf_files = glob.glob(os.path.join(base_path, '*.pdf'))
    if not pdf_files:
        print("Fehler: Keine PDF-Datei im Arbeitsverzeichnis gefunden.")
        print(f"Verzeichnis durchsucht: {base_path}")
        input("Drücken Sie Enter, um das Programm zu beenden...")
        raise FileNotFoundError('No PDF file found in the script directory.')
    pdf_path = pdf_files[0]
    print(f"PDF gefunden: {pdf_path}")
    csv_path = os.path.join(base_path, 'output.csv')

    data = []

    print("Starte PDF-Iteration...")
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                header1 = "Konto Bezeichnung Periode Belegdatum in PrCtrHW Refbeleg Kreditorbeschreibung Partnobjbe Materialbe Text Menge"
                header2 = "Per Objekt Objektbezeichnung Kostenart Menge Wert/KWähr Gegenkontobezeichnung Bezeichnung"
                # Split the text into blocks by headers
                blocks = []
                current_block = []
                current_header = None
                for line in text.split('\n'):
                    if line.startswith(header1):
                        if current_block:
                            blocks.append((current_header, current_block))
                        current_header = header1
                        current_block = [line]
                    elif line.startswith(header2):
                        if current_block:
                            blocks.append((current_header, current_block))
                        current_header = header2
                        current_block = [line]
                    elif not current_header and line.strip():
                        # Block without header
                        if current_block:
                            blocks.append((current_header, current_block))
                        current_header = None
                        current_block = [line]
                    else:
                        current_block.append(line)
                if current_block:
                    blocks.append((current_header, current_block))

                for block_header, block_lines in blocks:
                    if block_header == header2:
                        # --- original code for header2 block ---
                        for line in block_lines:
                            if line.replace(",", " ").replace(";", " ").startswith(header2):
                                continue  # skip header line
                            parts = line.split()
                            if not parts or not parts[0].isdigit():
                                continue  # skip lines that don't start with a number (Per)
                            per = parts[0]
                            # Objekt: next part, must have 5 segments separated by '.'
                            objekt = ''
                            objekt_idx = 1
                            if objekt_idx < len(parts) and parts[objekt_idx].count('.') == 4:
                                objekt = parts[objekt_idx]
                            else:
                                objekt = ''
                            # Objektbezeichnung: starts after Objekt, 5 segments, first separated by '-', rest by spaces, ends before a number (Kostenart)
                            ob_start = objekt_idx + 1
                            ob_end = ob_start
                            while ob_end < len(parts) and not parts[ob_end].replace('.', '', 1).isdigit():
                                ob_end += 1
                            objektbezeichnung = ' '.join(parts[ob_start:ob_end])
                            # Kostenart: first part that is exactly 6 digits, no separators
                            import re
                            kostenart = ''
                            kostenart_idx = None
                            for idx in range(ob_end, len(parts)):
                                if re.fullmatch(r"\d{6}", parts[idx]):
                                    kostenart = parts[idx]
                                    kostenart_idx = idx
                                    break
                            if kostenart_idx is not None:
                                ob_end = kostenart_idx + 1
                            # Menge: Zahl mit Komma und 3 Ziffern am Ende, optionalem Minus
                            menge = ''
                            if ob_end < len(parts) and re.fullmatch(r"\d+,\d{3}-?", parts[ob_end]):
                                menge = parts[ob_end]
                                ob_end += 1
                            # Wert/KWähr: Zahl mit Komma und 2 Ziffern am Ende, optionalem Minus
                            wert_kw = ''
                            if ob_end < len(parts) and re.fullmatch(r"\d+[.,]?\d*,\d{2}-?", parts[ob_end]):
                                wert_kw = parts[ob_end]
                                ob_end += 1
                            # Rest: Gegenkontobezeichnung und Bezeichnung zusammen
                            gegen_bez = ' '.join(parts[ob_end:]) if ob_end < len(parts) else ''
                            columns = [per, objekt, objektbezeichnung, kostenart, menge, wert_kw, gegen_bez]
                            data.append(columns)
                    elif block_header == header1:
                        # TODO: Add parsing logic for header1 block if needed
                        pass
                    else:
                        # TODO: Add parsing logic for blocks without header if needed
                        pass

    print("PDF-Iteration abgeschlossen.")
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerows(data)
    print(f"Daten extrahiert von {pdf_path} und gespeichert in {csv_path}")
    print("Programm erfolgreich abgeschlossen.")
except FileNotFoundError as e:
    print(f"FileNotFoundError: {e}")
    input("Drücken Sie Enter, um das Programm zu beenden...")
except Exception as e:
    print(f"Ein Fehler ist aufgetreten: {e}")
    input("Drücken Sie Enter, um das Programm zu beenden...")