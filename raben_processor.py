import pandas as pd
import os
import sys
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
    Řeší:
      - Desetinné čárky (12,5 -> 12.5)
      - Mezery v tisících (1 000 -> 1000)
      - Pevné mezery (\xa0)
      - Prázdné hodnoty
    """
    if pd.isna(val):
        return 0
    
    # Převedeme na string
    s = str(val).strip()
    
    # Pokud je to prázdný string
    if not s or s.lower() == 'nan':
        return 0
        
    try:
        # 1. Nahradíme desetinnou čárku tečkou
        s = s.replace(',', '.')
        # 2. Odstraníme běžné mezery i "tvrdé" mezery (non-breaking space)
        s = s.replace(' ', '').replace('\xa0', '')
        
        return float(s)
    except ValueError:
        # Pokud se převod nepovede, vrátíme 0 (nebo bychom mohli logovat chybu)
        return 0

def process_raben_file(input_path, output_path):
    print(f"--- Zpracovávám RABEN soubor: {input_path} ---")
    
    try:
        # 1. Autodetekce (z utils)
        sheet_name = find_best_sheet(input_path)
        # Poznámka: dtype=str zajistí, že se Excel nesnaží být chytrý při načítání,
        # my si typy vyřešíme sami ručně a bezpečněji.
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
            
        # Numerický sloupec - Použití naší robustní čistící funkce
        # Aplikujeme funkci clean_quantity na každý řádek ve sloupci
        print("   -> Provádím převod množství (oprava čárek/mezer)...")
        df["Mnozstvi_RABEN"] = df["Mnozstvi_RABEN"].apply(clean_quantity)
        
        # Kontrolní výpis pro debug (ukáže prvních 5 nenulových hodnot)
        non_zero = df[df["Mnozstvi_RABEN"] != 0]["Mnozstvi_RABEN"].head()
        if not non_zero.empty:
            print(f"   -> Ukázka načtených dat: {list(non_zero)}")
        else:
            print("   -> ⚠️ POZOR: Všechna data jsou 0, zkontroluj formát!")
        
        # 6. Export dat
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='RABEN', index=False)
            
        # 7. Formátování tabulky (z utils)
        create_excel_table(output_path, 'RABEN', 'tbl_RABEN')
        
        print(f"✅ Hotovo. RABEN uložen do: {output_path}")

    except Exception as e:
        print(f"❌ Chyba při zpracování RABEN: {e}")
        # Pro lepší debugování vypíšeme detail chyby
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