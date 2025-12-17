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

        os.makedirs(os.path.dirname(OUTPUT_MERGED), exist_ok=True)

        print("Zapisuji data...")
        with pd.ExcelWriter(OUTPUT_MERGED, engine='openpyxl') as writer:
            df_sap.to_excel(writer, sheet_name='SAP', index=False)
            df_raben.to_excel(writer, sheet_name='RABEN', index=False)

        # Formátování tabulek (volání centralizované funkce)
        print("Formátuji tabulky...")
        create_excel_table(OUTPUT_MERGED, "SAP", "tbl_SAP")
        create_excel_table(OUTPUT_MERGED, "RABEN", "tbl_RABEN")

        print(f"✅ Hotovo. Sloučený soubor uložen: {OUTPUT_MERGED}")

    except Exception as e:
        print(f"❌ Chyba při slučování: {e}")
        sys.exit(1)

if __name__ == "__main__":
    merge_files()