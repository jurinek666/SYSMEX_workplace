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

def process_raben_file(input_path, output_path):
    print(f"--- Zpracovávám RABEN soubor: {input_path} ---")
    
    try:
        # 1. Autodetekce (z utils)
        sheet_name = find_best_sheet(input_path)
        df = pd.read_excel(input_path, sheet_name=sheet_name)
        
        # 2. Očištění názvů sloupců
        df.columns = [c.strip() for c in df.columns]
        
        # 3. Kontrola a přejmenování
        missing_cols = [col for col in COLUMN_MAPPING.keys() if col not in df.columns]
        if missing_cols:
            raise ValueError(f"V souboru chybí sloupce: {', '.join(missing_cols)}")
            
        df = df.rename(columns=COLUMN_MAPPING)
        
        # 4. Výběr a uspořádání
        df = df[FINAL_ORDER]
        
        # 5. Úprava datových typů
        for col in ["Material", "Nazev", "Batch"]:
            df[col] = df[col].astype(str).replace('nan', '').str.strip()
            
        df["Mnozstvi_RABEN"] = pd.to_numeric(df["Mnozstvi_RABEN"], errors='coerce').fillna(0)
        
        # 6. Export dat
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='RABEN', index=False)
            
        # 7. Formátování tabulky (z utils)
        create_excel_table(output_path, 'RABEN', 'tbl_RABEN')
        
        print(f"✅ Hotovo. RABEN uložen do: {output_path}")

    except Exception as e:
        print(f"❌ Chyba při zpracování RABEN: {e}")
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