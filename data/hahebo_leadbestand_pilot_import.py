"""
Importeert de 5 pilotleads (L0001 t/m L0005) uit hahebo_pilot_leads.csv
in een kopie van HAHEBO_leadbestand_template.xlsx.

Werkwijze:
- Laadt het bestaande template (wordt zelf niet gewijzigd).
- Verwijdert alleen de fictieve TEST001-voorbeeldregel.
- Schrijft de 5 CSV-leads ongewijzigd (geen interpretatie, geen nieuwe
  waarden) in de bestaande 26 kolommen.
- Behoudt dropdowns, autofilter, bevroren rij, kolombreedtes, opmaak
  en het tabblad Uitleg van het template.
- Slaat het resultaat op als HAHEBO_leadbestand_pilot.xlsx.

Geen scraping, geen zoekfunctionaliteit, geen nieuwe classificatielogica.
"""

import csv
import os
from datetime import datetime
from openpyxl import load_workbook

DATE_COLUMN = "Datum verzameld"

BASE_DIR = os.path.dirname(__file__)
TEMPLATE_PATH = os.path.join(BASE_DIR, "HAHEBO_leadbestand_template.xlsx")
CSV_PATH = os.path.join(BASE_DIR, "hahebo_pilot_leads.csv")
OUTPUT_PATH = os.path.join(BASE_DIR, "HAHEBO_leadbestand_pilot.xlsx")

SHEET_NAME = "Leadbestand"


def load_pilot_rows():
    with open(CSV_PATH, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        return list(reader)


def main():
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb[SHEET_NAME]

    headers = [c.value for c in ws[1]]
    col_index = {name: i + 1 for i, name in enumerate(headers)}

    rows = load_pilot_rows()

    # Verwijder de fictieve TEST001-regel (rij 2): alle celwaarden leegmaken,
    # opmaak van de cellen blijft ongemoeid.
    for c in range(1, len(headers) + 1):
        ws.cell(row=2, column=c).value = None

    # Lead ID moet net als KVK-nummer/Postcode/Telefoonnummer als tekst
    # behandeld blijven, zodat Excel de waarde nooit herinterpreteert.
    lead_id_col = col_index["Lead ID"]
    for r in range(2, 501):
        ws.cell(row=r, column=lead_id_col).number_format = "@"

    # Schrijf de 5 pilotleads ongewijzigd, exact zoals in de CSV, vanaf rij 2.
    # "Datum verzameld" wordt als echte Excel-datum opgeslagen (i.p.v. tekst),
    # zodat later correct op datum gesorteerd en gefilterd kan worden. De
    # zichtbare notatie blijft dd-mm-jjjj via het bestaande celformaat.
    for r, row in enumerate(rows, start=2):
        for name in headers:
            raw = row.get(name, "")
            if name == DATE_COLUMN and raw != "":
                value = datetime.strptime(raw, "%d-%m-%Y").date()
            else:
                value = raw if raw != "" else None
            cell = ws.cell(row=r, column=col_index[name], value=value)
            if name == DATE_COLUMN and raw != "":
                cell.number_format = "dd-mm-yyyy"

    wb.save(OUTPUT_PATH)
    print(f"Bestand opgeslagen: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
