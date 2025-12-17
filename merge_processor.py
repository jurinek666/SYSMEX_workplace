import pandas as pd
import os
import sys
from utils import create_excel_table

# --- KONFIGURACE ---
INPUT_SAP = "sklady_porovnani/output/SAP.xlsx"
INPUT_RABEN = "sklady_porovnani/output/RABEN.xlsx"
OUTPUT_MERGED = "sklady_porovnani/input/POROVNANI_SKLADU.xlsx"

def merge_files():
    print(f"--- Spouštím slučování do: {OUTPUT_MERGED} ---")

    if not os.path.exists(INPUT_SAP):
        print(f"❌ CHYBA: Chybí vstup: {INPUT_SAP}")
        sys.exit(1)
    if not os.path.exists(INPUT_RABEN):
        print(f"❌ CHYBA: Chybí vstup: {INPUT_RABEN}")
        sys.exit(1)

    try:
        print("Načítám SAP a RABEN...")
        df_sap = pd.read_excel(INPUT_SAP, sheet_name="SAP")
        df_raben = pd.read_excel(INPUT_RABEN, sheet_name="RABEN")

        # ==========================================
        # --- RABEN BUSINESS LOGIC (TRANSFORMACE) ---
        # ==========================================
        print("Aplikuji obchodní pravidla na data RABEN...")
        
        # 1. Filtrace obalů (P-kódy)
        # Regex: ^ = začátek, P = písmeno P, \d{3,4} = 3 nebo 4 číslice, $ = konec
        p_code_mask = df_raben['Material'].astype(str).str.match(r'^P\d{3,4}$')
        count_deleted = p_code_mask.sum()
        
        if count_deleted > 0:
            # Ponecháme jen ty, co NEODPOVÍDAJÍ masce (vlnovka ~ znamená negaci)
            df_raben = df_raben[~p_code_mask].copy()
            print(f"   -> Odstraněno {count_deleted} řádků obalového materiálu (P-kódy).")
        
        # 2. Přepočet měrné jednotky pro ZE001906
        # Najdeme řádky
        convert_mask = df_raben['Material'] == 'ZE001906'
        count_converted = convert_mask.sum()
        
        if count_converted > 0:
            # Ujistíme se, že pracujeme s čísly (pro jistotu)
            df_raben['Mnozstvi_RABEN'] = pd.to_numeric(df_raben['Mnozstvi_RABEN'], errors='coerce').fillna(0)
            
            # Provedeme násobení 50 pouze u vybraných řádků
            df_raben.loc[convert_mask, 'Mnozstvi_RABEN'] = df_raben.loc[convert_mask, 'Mnozstvi_RABEN'] * 50
            print(f"   -> Přepočteno {count_converted} řádků materiálu ZE001906 (ks -> balení * 50).")

        # ==========================================
        # --- KONEC TRANSFORMACE ---
        # ==========================================

        os.makedirs(os.path.dirname(OUTPUT_MERGED), exist_ok=True)

        print("Zapisuji data...")
        with pd.ExcelWriter(OUTPUT_MERGED, engine='openpyxl') as writer:
            df_sap.to_excel(writer, sheet_name='SAP', index=False)
            df_raben.to_excel(writer, sheet_name='RABEN', index=False)

        # Formátování tabulek (volání centralizované funkce z utils.py)
        print("Formátuji tabulky...")
        create_excel_table(OUTPUT_MERGED, "SAP", "tbl_SAP")
        create_excel_table(OUTPUT_MERGED, "RABEN", "tbl_RABEN")

        print(f"✅ Hotovo. Sloučený soubor uložen: {OUTPUT_MERGED}")

    except Exception as e:
        print(f"❌ Chyba při slučování: {e}")
        # Pro lepší debugování v GitHub Actions
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    merge_files()