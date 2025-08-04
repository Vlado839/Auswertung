#Hilfsfunktion für "gesamtpreis":

import openpyxl
from typing import Optional, List, Tuple


def extrahiere_gesamtpreis(
        sheet: openpyxl.worksheet.worksheet.Worksheet
) -> Optional[float]:
    """
    Findet in einem Arbeitsblatt den Begriff 'Gesamt' (case-insensitive)
    und liefert die erste gefundene Zahl in derselben Zeile als float.
    """
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and "gesamt" in cell.value.lower():
                for c in row:
                    if isinstance(c.value, (int, float)):
                        return float(c.value)
    return None