# DUO open onderwijsdata

Deze map bevat officiële DUO-bronbestanden en, waar aangegeven, een
afgeleide dataset. Dit document legt per bestand de herkomst vast.

## Officiële bestanden (ongewijzigd)

- `onderwijspersoneel_vo_personen_2011_2025.xlsx`
  Bron: https://duo.nl/open_onderwijsdata/images/01.-onderwijspersoneel-vo-in-personen-2011-2025.xlsx
- `vestigingen_vo.csv`
  Bron: https://www.duo.nl/open_onderwijsdata/images/02.-alle-vestigingen-vo.csv
- `vestigingen_sbo_so_vso.csv`
  Bron: https://duo.nl/open_onderwijsdata/images/09.-alle-vestigingen-sbo-so-en-vso.csv

## PO-onderwijspersoneel: afgeleid bestand

- Officiële bron:
  https://duo.nl/open_onderwijsdata/images/01.-onderwijspersoneel-po-in-personen-2011-2025.xlsx

- Oorspronkelijke dataset:
  Onderwijspersoneel PO in aantal personen 2011-2025

- Origineel DUO-bestand NIET opgeslagen in deze repository wegens bestandsgrootte/uploadbeperking.

- De bestanden met "compact" in de naam zijn AFGELEID van het officiële DUO-bestand:
  - `duo_po_brin_personen_2011_2025_compact.csv`
  - `duo_po_brin_personen_2011_2025_compact.xlsx` (1-op-1 conversie van de CSV naar een geldig xlsx-bestand, geen inhoudelijke wijziging)

- Afgeleide inhoud:
  werkblad 6 / instellingsniveau / BRIN-gerelateerde data

- Behouden velden:
  - ONDERWIJSTYPE
  - BEVOEGD GEZAG
  - INSTELLINGSCODE
  - PERSONEN 2011 t/m PERSONEN 2025

- Bovenschoolse regels zijn niet opgenomen in de compacte versie.

- Waarden uit de geselecteerde regels zijn niet inhoudelijk aangepast.

- De compacte versie mag worden gebruikt voor de HAHEBO onderwijslookup, maar mag NIET worden aangeduid als het originele officiële DUO-bestand.

## BRIN -> aantal personen lookup

- Het script `scripts/duo_onderwijs_lookup.py` voert de lookup uit op basis
  van de bestanden in deze map (alleen lezen, nooit gewijzigd).

- Voorbeeldcommando:
  ```
  python scripts/duo_onderwijs_lookup.py 17HN
  python scripts/duo_onderwijs_lookup.py 20GB
  python scripts/duo_onderwijs_lookup.py --selftest
  ```

- Segmentatie (50+ / nog te bepalen / onbekend) vindt plaats op
  BRIN-/instellingsniveau, niet op vestigingsniveau. Eén BRIN kan meerdere
  fysieke vestigingen omvatten; het personeelsaantal is dan het totaal voor
  de hele instelling.

- Een vestigingscode (bijv. `17HN00`) wordt NIET als BRIN-input
  geaccepteerd. Het script accepteert uitsluitend een code van exact 4
  tekens en geeft anders een duidelijke foutmelding.

- Een niet-numerieke actuele waarde (`*`, `**`, `<5`, lege cel, andere
  tekst) wordt NOOIT naar 0 geconverteerd. In dat geval is
  `aantal_personen = null` en `segment = onbekend`.

- Een historische (oudere) numerieke waarde wordt uitsluitend als context
  teruggegeven (`laatste_numerieke_waarde`/`_jaar`/`_peildatum`) en bepaalt
  nooit het segment van de actuele waarde.

- PO-personeelsdata komt uit de afgeleide compacte dataset
  (`duo_po_brin_personen_2011_2025_compact.xlsx`), niet uit het originele
  officiële DUO-bestand. De scriptoutput labelt dit expliciet als
  `personeelsbron_type = afgeleide_compacte_DUO_dataset`.
