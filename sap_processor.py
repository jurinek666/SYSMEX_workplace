import pandas as pd
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
import sys
import os
import argparse

# Konfigurace názvů sloupců (interní reference)
REQUIRED_COLS = {
    "material": "Material",
    "material description": "Material description",
    "batch": "Batch",
    "total quantity": "Total Quantity",
    "storage location": "Storage location"
}

# Volitelný sloupec
OPTIONAL_COLS = {
    "plant": "Plant"
}

def normalize_columns(df):
    """
    Přejmenuje sloupce v DataFrame na standardní formát (dle zadání),
    aby nám nezáleželo na velikosti písmen v originále.
    """
    # Vytvoříme mapu: {názvy_v_souboru_lowercase: původní_název}
    actual_cols_lower = {c.lower().strip(): c for c in df.columns}
    
    mapping = {}
    missing = []
    
    # Kontrola povinných sloupců
    for req_key, req_std_name in REQUIRED_COLS.items():
        if req_key in actual_cols_lower:
            mapping[actual_cols_lower[req_key]] = req_std_name
        else:
            missing.append(req_std_name)
            
    # Kontrola volitelných
    for opt_key, opt_std_name in OPTIONAL_COLS.items():
        if opt_key in actual_cols_lower:
            mapping[actual_cols_lower[opt_key]] = opt_std_name

    if missing:
        raise ValueError(f"Chybí povinné sloupce: {', '.join(missing)}")
        
    # Přejmenování sloupců v DF na naše standardní názvy
    df = df.rename(columns=mapping)
    return df

def find_best_sheet(file_path):
    """
    Najde list s největší datovou plochou.
    """
    xl = pd.ExcelFile(file_path)
    best_sheet = None
    max_area = -1
    
    for sheet_name in xl.sheet_names:
        # Přečteme jen kousek pro rychlou kontrolu, ale musíme zjistit rozměry
        # Pandas nemá přímou metodu na 'rozměry bez načtení', načteme celý list
        # Pro velké soubory by se to dalo optimalizovat, pro běžné SAP exporty OK.
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        area = df.shape[0] * df.shape[1]
        
        # Musí mít hlavičku (neprázdný)
        if area > max_area and not df.empty:
            max_area = area
            best_sheet = sheet_name
            
    if best_sheet is None:
        raise ValueError("Nenašel jsem žádný list s daty.")
        
    print(f"Vybrán list pro zpracování: '{best_sheet}' (Plocha: {max_area} buněk)")
    return best_sheet

def process_file(input_path, output_path):
    print(f"--- Zpracovávám soubor: {input_path} ---")
    
    try:
        # 1. Autodetekce listu
        sheet_name = find_best_sheet(input_path)
        df = pd.read_excel(input_path, sheet_name=sheet_name)
        
        # 2. Normalizace a Validace sloupců
        df = normalize_columns(df)
        
        # 3. Transformace - Krok 1: Filtr Storage location
        # Ponech jen F010, F070
        allowed_locations = ["F010", "F070"]
        df = df[df["Storage location"].astype(str).isin(allowed_locations)].copy()
        
        # 4. Transformace - Krok 2: Smazání sloupců
        # Plant (pokud existuje) a Storage location
        cols_to_drop = ["Storage location"]
        if "Plant" in df.columns:
            cols_to_drop.append("Plant")
        df = df.drop(columns=cols_to_drop, errors='ignore')
        
        # 5. Transformace - Krok 3: Uspořádání sloupců
        # Požadované pořadí: Material, Material description, Batch, Total Quantity
        ordered_cols = ["Material", "Material description", "Batch", "Total Quantity"]
        # Pokud tam jsou nějaké navíc (což by neměly být, ale pro jistotu), dáme je na konec
        remaining_cols = [c for c in df.columns if c not in ordered_cols]
        df = df[ordered_cols + remaining_cols]
        
        # 6. Transformace - Krok 4: Přejmenování
        rename_map = {
            "Material description": "Název",
            "Total Quantity": "Mnozstvi_SAP"
        }
        df = df.rename(columns=rename_map)
        
        # 7. Transformace - Krok 5: Sort
        # Převod na číslo, NaN -> 0
        df["Mnozstvi_SAP"] = pd.to_numeric(df["Mnozstvi_SAP"], errors='coerce').fillna(0)
        # Sort sestupně
        df = df.sort_values(by="Mnozstvi_SAP", ascending=False, kind='mergesort')
        
        # 8. Transformace - Krok 6: Smazat první datový řádek po sortu
        if len(df) > 0:
            df = df.iloc[1:] # Smaže řádek s indexem 0 (první řádek dat)
        
        # 9. Transformace - Krok 7: Ořezat na 4 sloupce
        df = df.iloc[:, :4]
        
        # --- EXPORT DO EXCELU S TABULKOU ---
        # Použijeme Pandas pro zápis dat, ale OpenPyXL pro formátování tabulky
        
        # Zápis dat
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='SAP', index=False)
            
        # Otevření pro formátování
        wb = openpyxl.load_workbook(output_path)
        ws = wb["SAP"]
        
        # Definice rozsahu tabulky A1:D{last_row}
        max_row = ws.max_row
        max_col = 4 # A-D
        # Pokud je tabulka prázdná (jen hlavička), max_row je 1
        ref = f"A1:D{max_row}"
        
        # Vytvoření tabulky "tbl_SAP"
        tab = Table(displayName="tbl_SAP", ref=ref)
        
        # Styl tabulky (modrý standardní)
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        tab.tableStyleInfo = style
        
        ws.add_table(tab)
        wb.save(output_path)
        
        print(f"✅ Hotovo. Uloženo do: {output_path}")

    except Exception as e:
        print(f"❌ Chyba při zpracování {input_path}: {e}")
        # V reálném nasazení bychom zde mohli ukončit sys.exit(1), 
        # ale pokud jedeme smyčku přes více souborů, chceme pokračovat.
        if mode == 'single':
            sys.exit(1)

# --- CLI Orchestrace ---
if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", choices=["all", "single"], default="all")
    parser.add_argument("--file_path", help="Cesta k souboru pro mode=single")
    args = parser.parse_args()
    
    mode = args.mode
    input_dir = "sklady_porovnani/input"
    output_dir = "sklady_porovnani/output"
    
    os.makedirs(output_dir, exist_ok=True)

    if mode == "single":
        if not args.file_path:
            print("Chyba: Pro mode=single musíš zadat --file_path")
            sys.exit(1)
        
        # Ošetření cesty
        infile = args.file_path
        filename = os.path.basename(infile)
        outfile = os.path.join(output_dir, filename)
        process_file(infile, outfile)

    elif mode == "all":
        if not os.path.exists(input_dir):
             print(f"Složka {input_dir} neexistuje.")
             sys.exit(1)
             
        for filename in os.listdir(input_dir):
            if filename.lower().endswith(".xlsx") and not filename.startswith("~$"):
                infile = os.path.join(input_dir, filename)
                outfile = os.path.join(output_dir, filename)
                process_file(infile, outfile)