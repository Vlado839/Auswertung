import os
import pandas as pd
import time
from analyse import analyse_angebote
from styling import style_auswertungs_sheet, style_pivot_sheet


def main():
    print("🚀 Starte Angebotsauswertung...\n")

    # Pfade
    # Zadaj cestu k priečinku, kde sú .xlsx súbory s ponukami
    angebote_dir = r"C:\Users\A102569436\OneDrive - Deutsche Telekom AG\Doc\Privat\Python\Angebote"

    # Zadaj cestu k súboru banf_volumen.xlsx
    banf_path = r"C:\Users\A102569436\OneDrive - Deutsche Telekom AG\Doc\Privat\Python\Angebote\banf_volumen.xlsx"

    # Zadaj cestu k výstupnému súboru (bude prepísaný)
    output_path = r"C:\Users\A102569436\OneDrive - Deutsche Telekom AG\Doc\Privat\Python\Angebote\Auswertung_Angebote.xlsx"

    # Input-Prüfung
    if not os.path.exists(angebote_dir):
        print(f"❌ Ordner '{angebote_dir}' fehlt!")
        return
    if not os.path.isfile(banf_path):
        print(f"❌ Datei '{banf_path}' fehlt!")
        return

    print("🔍 Analysiere Angebote...")
    auswertung = analyse_angebote(angebote_dir, banf_path)
    if not auswertung:
        print("⚠️ Keine gültigen Daten gefunden.")
        return

    df = pd.DataFrame(auswertung).sort_values(by=["Los", "Gesamtpreis (€)"])

    # Pivot-Sheet für Vergleich: Los = Zeile, Lieferant = Spalte, Wert = Preis
    pivot_df = df.pivot_table(
        index="Los",
        columns="Lieferant",
        values="Gesamtpreis (€)",
        aggfunc="min"
    ).round(2).reset_index()

    # Output-Verzeichnis vorbereiten
    os.makedirs("output", exist_ok=True)

    print("💾 Schreibe Auswertung in Excel...")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Sheet: Übersicht (alle Angebote)
        df.to_excel(writer, index=False, sheet_name="Übersicht")
        style_auswertungs_sheet(writer.sheets["Übersicht"], list(df.columns))

        # Sheet: Vergleich (pivotiert mit Bestbieter)
        pivot_df.to_excel(writer, index=False, sheet_name="Vergleich")
        style_pivot_sheet(writer.sheets["Vergleich"], list(pivot_df.columns))

        # Sheet pro Los
        for los_name in sorted(df["Los"].dropna().unique()):
            los_df = df[df["Los"] == los_name].copy()
            sheet_name = f"Los_{los_name}"[:31]  # Excel-Sheetnamen max. 31 Zeichen
            los_df.to_excel(writer, index=False, sheet_name=sheet_name)
            style_auswertungs_sheet(writer.sheets[sheet_name], list(los_df.columns))

    time.sleep(1)
    if os.path.exists(output_path):
        print("📂 Öffne Auswertung...")
        os.startfile(output_path)
    else:
        print("❌ Fehler beim Schreiben der Datei.")


if __name__ == "__main__":
    main()