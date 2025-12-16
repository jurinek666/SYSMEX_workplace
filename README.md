# Automatická úprava SAP Excelu (GitHub Actions)

Tento repo automaticky zpracuje Excel soubory nahrané do `sklady_porovnani/input/` a vyrobí vyčištěný výstup:
- filtr `Storage location` ∈ {`F010`,`F070`} → ostatní řádky smaže
- smaže sloupce `Plant` a `Storage location`
- pořadí sloupců: `Material`, `Material description`, `Batch`, `Total Quantity`
- přejmenuje `Material description` → `Název`, `Total Quantity` → `Mnozstvi_SAP`
- seřadí podle `Mnozstvi_SAP` (sestupně), odstraní první datový řádek
- ponechá pouze první 4 sloupce; uloží list `SAP` s tabulkou `tbl_SAP` (A:D)

Výsledek stáhneš z GitHub Actions jako **artifact** `SAP-outputs`.

---

## Struktura repozitáře

/
├─ .github/
│ └─ workflows/
│ └─ transform-sap.yml # workflow (spouští Python skript)
├─ sklady_porovnani/
│ ├─ requirements.txt # závislosti (pandas, openpyxl)
│ ├─ scripts/
│ │ └─ transform_sap.py # hlavní transformační skript
│ ├─ input/ # sem nahraj zdrojové .xlsx
│ │ └─ .gitkeep
│ └─ output/ # lokální výstupy (ignorovány v gitu)
│ └─ .gitkeep
└─ .gitignore


> Workflow se spustí automaticky při pushi změn v:
> - `sklady_porovnani/input/**/*.xlsx`
> - `sklady_porovnani/scripts/**`
> - `sklady_porovnani/requirements.txt`
> - `.github/workflows/transform-sap.yml`

---

## Jak to použít (bez terminálu)

1. **Vytvoř složky/soubory**  
   - V GitHubu klikni: **Add file → Create new file**  
   - Založ postupně:
     - `.github/workflows/transform-sap.yml`
     - `sklady_porovnani/requirements.txt`
     - `sklady_porovnani/scripts/transform_sap.py`
     - `sklady_porovnani/input/.gitkeep`
     - `sklady_porovnani/output/.gitkeep`
   - Vlož připravené obsahy ze zadání (nebo z PR).

2. **Nahraj Excel**  
   - **Add file → Upload files**  
   - Cílová složka: `sklady_porovnani/input/`  
   - Nahraj libovolný `.xlsx` se správnou hlavičkou (řádek 1) a názvy sloupců.

3. **Zkontroluj běh**  
   - Repozitář → záložka **Actions** → workflow **Transform SAP Excel**  
   - Po doběhu otevři běh → **Artifacts → SAP-outputs** → stáhni ZIP.  
   - Uvnitř najdeš `*_SAP.xlsx`.

> Tip: můžeš nahrát víc `.xlsx` najednou – workflow je zpracuje všechny.

---

## Lokální běh (volitelné)

> Vyžaduje Python 3.11+

```bash
python -m venv .venv
. .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r sklady_porovnani/requirements.txt
python sklady_porovnani/scripts/transform_sap.py \
  --input sklady_porovnani/input/TVUJ_SOUBOR.xlsx \
  --output sklady_porovnani/output/TVUJ_SOUBOR_SAP.xlsx
