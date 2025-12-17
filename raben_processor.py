import pandas as pd
import os
import sys
# Předpokládáme, že utils.py existuje ve stejné složce
from utils import find_best_sheet, create_excel_table

# --- KONFIGURACE ---
COLUMN_MAPPING = {
    "1-Císlo zboží": "Material",
    "3-název": "Nazev",
    "12-šarže": "Batch",
    "4-ks": "Mnozstvi_RABEN"
}

FINAL_ORDER = ["Material", "Nazev", "Batch", "Mnozstvi_RABEN"]

def clean_quantity(val):
    """
    Pomocná funkce pro bezpečný převod na číslo (float).
    Řeší formáty: "1 500,50", "1.500,50", "1500,50".
    """
    if pd.isna(val):
        return 0
    
    # Převedeme na string
    s = str(val).strip()
    
    # Pokud je to prázdný string nebo 'nan'
    if not s or s.lower() == 'nan':
        return 0
        
    try:
        # 1. Odstraníme mezery (běžné i tvrdé/non-breaking)
        s = s.replace(' ', '').replace('\xa0', '')
        
        # 2. Logika pro odstranění oddělovačů tisíců (teček)
        # Pokud řetězec obsahuje čárku I tečku (např. "1.200,50"), tečka je oddělovač tisíců -> pryč s ní.
        if ',' in s and '.' in s:
            s = s.replace('.', '')
            
        # 3. Nahradíme desetinnou čárku tečkou (aby tomu rozuměl Python float)
        s = s.replace(',', '.')
        
        return float(s)
    except ValueError:
        # Debug výpis, abychom v logu viděli, na čem to spadlo (pokud by se to stalo znovu)
        print(f"⚠️ Varování: Hodnotu '{val}' nelze převést na číslo. Nahrazuji 0.")
        return 0

def process_raben_file(input_path, output_path):
    print(f"--- Zpracovávám RABEN soubor: {input_path} ---")
    
    try:
        # 1. Autodetekce (z utils)
        sheet_name = find_best_sheet(input_path)
        
        # dtype=str zajistí, že načteme "raw" data a Excel neudělá nechtěné konverze
        df = pd.read_excel(input_path, sheet_name=sheet_name, dtype=str)
        
        # 2. Očištění názvů sloupců
        df.columns = [str(c).strip() for c in df.columns]
        
        # 3. Kontrola a přejmenování
        missing_cols = [col for col in COLUMN_MAPPING.keys() if col not in df.columns]
        if missing_cols:
            raise ValueError(f"V souboru chybí sloupce: {', '.join(missing_cols)}")
            
        df = df.rename(columns=COLUMN_MAPPING)
        
        # 4. Výběr a uspořádání
        df = df[FINAL_ORDER]
        
        # 5. Úprava datových typů (OPRAVENO)
        
        # Textové sloupce
        for col in ["Material", "Nazev", "Batch"]:
            df[col] = df[col].astype(str).replace('nan', '').str.strip()
            
        # Numerický sloupec - Aplikace naší vylepšené funkce clean_quantity
        print("   -> Provádím převod množství (oprava formátu čísel)...")
        df["Mnozstvi_RABEN"] = df["Mnozstvi_RABEN"].apply(clean_quantity)
        
        # Kontrolní výpis pro jistotu (zobrazí součet, abychom viděli, že to není 0)
        total_qty = df["Mnozstvi_RABEN"].sum()
        print(f"   -> Kontrola: Celkový součet množství je {total_qty}")
        
        # 6. Export dat
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='RABEN', index=False)
            
        # 7. Formátování tabulky (z utils)
        create_excel_table(output_path, 'RABEN', 'tbl_RABEN')
        
        print(f"✅ Hotovo. RABEN uložen do: {output_path}")

    except Exception as e:
        print(f"❌ Chyba při zpracování RABEN: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    input_dir = "sklady_porovnani/input"
    output_dir = "sklady_porovnani/output"
    filename = "RABEN.xlsx"
    
    infile = os.path.join(input_dir, filename)
    outfile = os.path.join(output_dir, filename)
    
    os.makedirs(output_dir, exist_ok=True)

    if not os.path.exists(infile):
        print(f"❌ CHYBA: Soubor '{infile}' neexistuje.")
        sys.exit(1)

    process_raben_file(infile, outfile)