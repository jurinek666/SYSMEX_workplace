import pandas as pd
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
import os
import sys

# --- KONFIGURACE ---
# Mapování: { "Původní název v Excelu": "Náš cílový název" }
COLUMN_MAPPING = {
    "1-Císlo zboží": "Material",
    "3-název": "Nazev",
    "12-šarže": "Batch",
    "4-ks": "Mnozstvi_RABEN"
}

# Požadované pořadí sloupců ve výstupu
FINAL_ORDER = ["Material", "Nazev", "Batch", "Mnozstvi_RABEN"]

# --- FUNKCE ---

def find_best_sheet(file_path):
    """
    Najde list s největší datovou plochou (počet řádků * počet sloupců).
    Slouží k automatické detekci správného listu, pokud jich je více.
    """
    try:
        xl = pd.ExcelFile(file_path)
    except Exception as e:
        raise ValueError(f"Nelze otevřít Excel soubor: {e}")

    best_sheet = None
    max_area = -1
    
    for sheet_name in xl.sheet_names:
        # Načteme list
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Spočítáme plochu (ignorujeme prázdné listy)
        area = df.shape[0] * df.shape[1]
        
        if area > max_area and not df.empty:
            max_area = area
            best_sheet = sheet_name
            
    if best_sheet is None:
        raise ValueError("Nenašel jsem žádný list s daty.")
        
    print(f"Vybrán list pro zpracování: '{best_sheet}' (Plocha: {max_area} buněk)")
    return best_sheet

def process_raben_file(input_path, output_path):
    print(f"--- Zpracovávám RABEN soubor: {input_path} ---")
    
    try:
        # 1. Autodetekce a načtení listu
        sheet_name = find_best_sheet(input_path)
        df = pd.read_excel(input_path, sheet_name=sheet_name)
        
        # 2. Očištění názvů sloupců (strip whitespace)
        # Odstraní mezery na začátku a konci názvů sloupců v původním souboru
        df.columns = [c.strip() for c in df.columns]
        
        # 3. Kontrola a přejmenování sloupců
        # Zkontrolujeme, zda existují všechny potřebné sloupce
        missing_cols = [col for col in COLUMN_MAPPING.keys() if col not in df.columns]
        if missing_cols:
            raise ValueError(f"V souboru chybí tyto sloupce: {', '.join(missing_cols)}")
            
        # Přejmenování
        df = df.rename(columns=COLUMN_MAPPING)
        
        # 4. Výběr a uspořádání sloupců
        # Vybereme jen ty, co nás zajímají, a seřadíme je
        df = df[FINAL_ORDER]
        
        # 5. Úprava datových typů
        
        # Textové sloupce (Material, Nazev, Batch)
        # Převedeme na string, NaN nahradíme prázdným řetězcem
        text_cols = ["Material", "Nazev", "Batch"]
        for col in text_cols:
            df[col] = df[col].astype(str).replace('nan', '').str.strip()
            
        # Numerický sloupec (Mnozstvi_RABEN)
        # Coerce: nečíselné hodnoty se změní na NaN. Fillna(0): NaN se změní na 0.
        df["Mnozstvi_RABEN"] = pd.to_numeric(df["Mnozstvi_RABEN"], errors='coerce').fillna(0)
        
        # 6. Export do Excelu (zatím bez tabulky)
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='RABEN', index=False)
            
        # 7. Formátování tabulky (openpyxl)
        wb = openpyxl.load_workbook(output_path)
        ws = wb["RABEN"]
        
        last_row = ws.max_row
        # Rozsah A1:D{last_row}
        ref = f"A1:D{last_row}"
        
        # Vytvoření tabulky
        tab = Table(displayName="tbl_RABEN", ref=ref)
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        tab.tableStyleInfo = style
        
        ws.add_table(tab)
        wb.save(output_path)
        
        print(f"✅ Hotovo. RABEN uložen do: {output_path}")

    except Exception as e:
        print(f"❌ Chyba při zpracování RABEN: {e}")
        sys.exit(1)

# --- SPUŠTĚNÍ ---
if __name__ == "__main__":
    # Definice cest
    input_dir = "sklady_porovnani/input"
    output_dir = "sklady_porovnani/output"
    filename = "RABEN.xlsx"
    
    infile = os.path.join(input_dir, filename)
    outfile = os.path.join(output_dir, filename)
    
    # Zajištění výstupní složky
    os.makedirs(output_dir, exist_ok=True)

    # Kontrola existence vstupu
    if not os.path.exists(infile):
        print(f"❌ CHYBA: Soubor '{infile}' neexistuje.")
        print("Nahraj prosím soubor 'RABEN.xlsx' do složky 'sklady_porovnani/input/'.")
        sys.exit(1)

    # Spuštění procesu
    process_raben_file(infile, outfile)