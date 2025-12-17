import pandas as pd
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
import os
import sys

# --- KONFIGURACE ---
INPUT_SAP = "sklady_porovnani/output/SAP.xlsx"
INPUT_RABEN = "sklady_porovnani/output/RABEN.xlsx"
OUTPUT_MERGED = "sklady_porovnani/input/POROVNANI_SKLADU.xlsx"

def add_table_formatting(ws, table_name):
    """
    Pomocná funkce, která na daném listu (ws) vytvoří Excel tabulku
    přes celý rozsah dat.
    """
    max_row = ws.max_row
    max_col = ws.max_column
    
    # Pokud je list prázdný nebo má jen hlavičku, nic neděláme (nebo jen hlavičku)
    if max_row < 1:
        return

    # Získáme písmeno posledního sloupce (např. 4 -> D)
    last_col_letter = openpyxl.utils.get_column_letter(max_col)
    
    # Definice rozsahu např. "A1:D150"
    ref = f"A1:{last_col_letter}{max_row}"
    
    # Vytvoření tabulky
    tab = Table(displayName=table_name, ref=ref)
    
    # Styl (modrý pruhovaný)
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    tab.tableStyleInfo = style
    
    ws.add_table(tab)

def merge_files():
    print(f"--- Spouštím slučování do: {OUTPUT_MERGED} ---")

    # 1. Kontrola existence vstupů
    if not os.path.exists(INPUT_SAP):
        print(f"❌ CHYBA: Chybí vstupní soubor: {INPUT_SAP}")
        sys.exit(1)
    if not os.path.exists(INPUT_RABEN):
        print(f"❌ CHYBA: Chybí vstupní soubor: {INPUT_RABEN}")
        sys.exit(1)

    try:
        # 2. Načtení dat
        print("Načítám SAP a RABEN...")
        df_sap = pd.read_excel(INPUT_SAP, sheet_name="SAP")
        df_raben = pd.read_excel(INPUT_RABEN, sheet_name="RABEN")

        # Zajištění existence cílové složky (input)
        os.makedirs(os.path.dirname(OUTPUT_MERGED), exist_ok=True)

        # 3. Zápis dat do nového sešitu
        print("Zapisuji data...")
        with pd.ExcelWriter(OUTPUT_MERGED, engine='openpyxl') as writer:
            df_sap.to_excel(writer, sheet_name='SAP', index=False)
            df_raben.to_excel(writer, sheet_name='RABEN', index=False)

        # 4. Formátování tabulek (OpenPyXL)
        print("Formátuji tabulky...")
        wb = openpyxl.load_workbook(OUTPUT_MERGED)
        
        # Formátování listu SAP
        if "SAP" in wb.sheetnames:
            add_table_formatting(wb["SAP"], "tbl_SAP")
            
        # Formátování listu RABEN
        if "RABEN" in wb.sheetnames:
            add_table_formatting(wb["RABEN"], "tbl_RABEN")

        # Uložení
        wb.save(OUTPUT_MERGED)
        print(f"✅ Hotovo. Sloučený soubor uložen: {OUTPUT_MERGED}")

    except Exception as e:
        print(f"❌ Chyba při slučování: {e}")
        sys.exit(1)

if __name__ == "__main__":
    merge_files()