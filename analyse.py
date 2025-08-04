import os
import glob
import pandas as pd
import openpyxl
from typing import List, Dict, Any
from utils import extrahiere_gesamtpreis

SUPPLIER_CELL = (12, 4)  # D12

UNGUELTIGE_NAMEN = {"xxx GmbH", "Firma", "test", "n/a", "-", "", None}


def analyse_angebote(
        angebote_dir: str,
        banf_path: str
) -> List[Dict[str, Any]]:
    records: List[Dict[str, Any]] = []

    for path in glob.glob(os.path.join(angebote_dir, '*.xlsx')):
        wb = openpyxl.load_workbook(path, data_only=True)
        angebot_eintraege = []

        # Schritt 1: Alle Sheets einlesen und Rohdaten sammeln
        for sheet in wb.worksheets:
            raw_supplier = sheet.cell(*SUPPLIER_CELL).value
            raw_supplier = str(raw_supplier).strip() if raw_supplier else ""

            total = extrahiere_gesamtpreis(sheet)

            angebot_eintraege.append({
                'Lieferant': raw_supplier,
                'Los': sheet.title,
                'Gesamtpreis (€)': total
            })

        # Schritt 2: Einen gültigen Namen suchen
        gueltiger_name = next(
            (e['Lieferant'] for e in angebot_eintraege if e['Lieferant'] not in UNGUELTIGE_NAMEN),
            "Unbekannt"
        )

        # Schritt 3: Allen Einträgen denselben Namen geben (wenn nötig)
        for e in angebot_eintraege:
            if e['Lieferant'] in UNGUELTIGE_NAMEN:
                e['Lieferant'] = gueltiger_name
            records.append(e)

    # DataFrame erzeugen
    df = pd.DataFrame(records)
    df['Los'] = df['Los'].astype(str)

    # BANF einlesen
    try:
        banf = pd.read_excel(banf_path, dtype={'Los': str})
    except Exception:
        banf = pd.DataFrame()

    if 'Banf-Volumen (€)' in banf.columns:
        banf.rename(columns={'Banf-Volumen (€)': 'Banf (€)'}, inplace=True)

    if {'Los', 'Banf (€)'}.issubset(banf.columns):
        df = df.merge(banf[['Los', 'Banf (€)']], on='Los', how='left')
    else:
        df['Banf (€)'] = pd.NA

    # Abweichungen berechnen
    df['Abweichung (€)'] = df['Gesamtpreis (€)'] - df['Banf (€)']
    df['Abweichung (%)'] = df.apply(
        lambda r: (r['Abweichung (€)'] / r['Banf (€)'])
        if pd.notna(r['Banf (€)']) and r['Banf (€)'] != 0 else pd.NA,
        axis=1
    )

    return df.to_dict(orient='records')