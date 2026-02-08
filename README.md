# Automatisk vakansberäkning

Professionell lösning för att beräkna karens och OB-ersättning för vakanta sjukskift baserat på sjuklistor och lönebesked.

## 🚀 Snabbstart

### Installation

```bash
# Installera dependencies
pip install -r requirements.txt

# Kör web-appen
streamlit run vakant_karens_streamlit.py
```

Öppna sedan din webbläsare på `http://localhost:8501`

### Kommandorad (CLI)

```bash
python vakant_karens_app.py \
  --sick_pdf Sjuklista_december_2025.pdf \
  --payslips person1.pdf person2.pdf person3.pdf \
  --out rapport.xlsx
```

## 📁 Filstruktur

```
.
├── vakant_karens_app.py        # Huvudmodul med all logik
├── vakant_karens_streamlit.py  # Web-gränssnitt
├── requirements.txt             # Python-dependencies
└── README.md                    # Denna fil
```

## 🎯 Funktioner

### Förbättringar från original-versionen

✅ **Dynamisk PDF-parsing**: Automatisk detektion av sjuklistesidor (ingen hardkodning)
✅ **Robust felhantering**: Validering och tydliga felmeddelanden
✅ **Logging**: Spårning av vad som händer
✅ **Konfigurationsbar**: Externa inställningar för helgdagar
✅ **Web-gränssnitt**: Drag-and-drop uppladdning via Streamlit
✅ **Progress tracking**: Se vad som bearbetas
✅ **Modulär arkitektur**: Lätt att underhålla och utöka

### Huvudfunktioner

- **Karens-beräkning**: Korrekt förbrukning av karenssaldo över hela dagen
- **OB-klassificering**: Automatisk kategorisering (Helg-OB, Natt, Kväll, Dag)
- **GT14-hantering**: Särskild hantering för sjukperioder >14 dagar
- **Vakant-filtrering**: Visar endast segment där ersättare saknas
- **Detaljerade rapporter**: Excel med både detaljer och sammanfattningar

## 📊 Input-filer

### Sjuklista PDF
- Innehåller kolumner: "Sjukskriven" och "Vikarie"
- Automatisk detektion av sidnummer
- Format: `Sjuklista [månad] [år]`

Exempel:
```
Sjuklista december 2025

Datum    Tid         Timmar  Sjukskriven              Vikarie
25       08:00-16:00  8,0    Anna Andersson 199001011234  ...
```

### Lönebesked PDF
- Ett lönebesked per person
- Filnamn ska innehålla personnummer: `...-YYMMDD-XXXX.pdf`
- Innehåller:
  - Anställningsnr
  - Karens (löneart 43100/43101)
  - Sjuk dag >14 (löneart 433... dag 15--)

## 📈 Output

Excel-fil med följande flikar:

### 1. Detalj
Alla segment med kolumner:
- Anställningsnr, Personnummer, Namn
- Datum, Start, Slut, Timmar
- OB-klass (Helg-OB/Natt/Kväll/Dag)
- Status (Betald/Karens/Karens och >14)
- Betalda timmar (vakant)

### 2. Summering_Betald
Sammanställning per person och OB-klass för betalda timmar

### 3. Summering_Karens
Sammanställning för karens-timmar

### 4. Summering_>14
Sammanställning för sjukperioder över 14 dagar

### 5. Summering_UnderlagSaknas
Timmar där lönebesked saknas

## 🏷️ OB-klassificering

### Helg-OB
- Lördagar & söndagar: 00:00-24:00
- Helgdagar: 00:00-24:00
- Fredag & dag före helgdag: 19:00-24:00
- Måndag & dag efter helgdag: 00:00-07:00

### Natt
- 22:00-06:00 (vardagar)

### Kväll
- 19:00-22:00 (vardagar)

### Dag
- Övrig tid

## ⚙️ Konfiguration

### Helgdagar

Standard-helgdagar finns i koden, men kan läggas till via:

**Web-appen**: Använd sidebar för att lägga till extra helgdagar

**CLI**: Skapa en config-fil:

```python
from datetime import date
from vakant_karens_app import load_config

custom_holidays = [
    date(2026, 1, 6),   # Trettondagen
    date(2026, 6, 6),   # Nationaldagen
    # ... fler helgdagar
]

config = load_config(holidays=custom_holidays)
```

## 🔧 Avancerad användning

### Programmatisk integration

```python
from vakant_karens_app import process_karens_calculation

process_karens_calculation(
    sick_pdf="sjuklista.pdf",
    payslip_paths=["person1.pdf", "person2.pdf"],
    output_xlsx="rapport.xlsx"
)
```

### Custom logging

```python
import logging

# Sätt till DEBUG för detaljerad information
logging.getLogger().setLevel(logging.DEBUG)

# Eller skapa egen logger
logger = logging.getLogger("vakant_karens")
logger.setLevel(logging.INFO)
```

### Batch-processing

```python
from pathlib import Path
from vakant_karens_app import process_karens_calculation

# Hitta alla sjuklistor
sick_lists = Path("./sjuklistor").glob("Sjuklista*.pdf")

for sick_pdf in sick_lists:
    month = sick_pdf.stem.split("_")[1]
    output = f"rapport_{month}.xlsx"
    
    # Hitta matchande lönebesked
    payslips = list(Path("./lonebesked").glob(f"*{month}*.pdf"))
    
    process_karens_calculation(
        str(sick_pdf),
        [str(p) for p in payslips],
        output
    )
```

## 🏗️ Arkitektur

### Huvudklasser

- **Config**: Konfiguration och inställningar
- **SwedishDateHelper**: Svenska datum och helgdagslogik
- **OBClassifier**: Klassificerar tid till OB-kategori
- **PersonnummerParser**: Hanterar personnummer
- **PayslipParser**: Extraherar data från lönebesked
- **SickListParser**: Extraherar data från sjuklistor
- **KarensCalculator**: Beräknar karens och segmenterar
- **ReportGenerator**: Skapar Excel-rapporter

### Dataflöde

```
Sjuklista PDF + Lönebesked PDFs
           ↓
    Parse & Extract
           ↓
  Calculate Segments
  (OB + Karens logic)
           ↓
    Merge & Process
           ↓
    Excel Report
```

## 🐛 Felsökning

### "No sick leave data found"
- Kontrollera att PDF:en innehåller text (inte bara bilder)
- Sätt debug-läge: `--verbose` för att se vad som parsas

### "Could not extract personnummer"
- Filnamn måste innehålla: `YYMMDD-XXXX` format
- Exempel: `lonebesked-900101-1234.pdf`

### "Page X out of range"
- PDF:en har för få sidor
- Använd `--verbose` för att se vilka sidor som detekteras

### PDF-parsing ger fel data
- Kontrollera att PDF:en är text-baserad (inte skannad bild)
- Testa med `pdfplumber` direkt:
  ```python
  import pdfplumber
  with pdfplumber.open("fil.pdf") as pdf:
      print(pdf.pages[0].extract_text())
  ```

## 📝 Utveckling

### Lägga till nya funktioner

1. **Ny OB-kategori**: Uppdatera `OBClassifier.classify()`
2. **Ny löneart**: Lägg till i `Config.karens_codes`
3. **Nytt output-format**: Utöka `ReportGenerator`

### Testa manuellt

```bash
# Testa parsing av en sjuklista
python -c "
from vakant_karens_app import SickListParser, load_config
parser = SickListParser(load_config())
df = parser.parse_sick_rows('test.pdf')
print(df)
"

# Testa OB-klassificering
python -c "
from vakant_karens_app import OBClassifier, load_config
from datetime import datetime
classifier = OBClassifier(load_config().holidays)
print(classifier.classify(datetime(2025, 12, 25, 15, 0)))  # Helg-OB
print(classifier.classify(datetime(2025, 12, 23, 23, 0)))  # Natt
"
```

## 📄 Licens

Denna kod är skapad för intern användning.

## 🤝 Support

För frågor eller problem:
1. Kontrollera detta README
2. Kör med `--verbose` för detaljerad information
3. Kontrollera logs i terminalen

## 📚 Dependencies

- **pandas**: Datahantering och Excel-output
- **pdfplumber**: PDF-parsing (text extraction)
- **openpyxl**: Excel-filhantering
- **streamlit**: Web-gränssnitt (optional)

## 🔄 Versionshistorik

### v2.0 (Improved Version)
- ✅ Dynamisk PDF-detektion
- ✅ Modulär arkitektur
- ✅ Web-gränssnitt
- ✅ Förbättrad felhantering
- ✅ Logging och progress tracking
- ✅ Konfigurerbar

### v1.0 (Original)
- Grundläggande funktionalitet
- Hårdkodade sidnummer
- CLI-only
