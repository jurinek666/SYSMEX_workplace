import pandas as pd
import sys
import os
# Import vlastních funkcí
from utils import find_best_sheet, create_excel_table

# --- KONFIGURACE ---
REQUIRED_COLS = {
    "material": "Material",
    "material description": "Material description",
    "batch": "Batch",
    "total quantity": "Total Quantity",
    "storage location": "Storage location"
}

OPTIONAL_COLS = {
    "plant": "Plant"
}

def normalize_columns(df):
    """
    Přejmenuje sloupce v DataFrame na standardní formát (dle zadání).
    """
    actual_cols_lower = {c.lower().strip(): c for c in df.columns}
    mapping = {}
    missing = []
    
    for req_key, req_std_name in REQUIRED_COLS.items():
        if req_key in actual_cols_lower:
            mapping[actual_cols_lower[req_key]] = req_std_name
        else:
            missing.append(req_std_name)
            
    for opt_key, opt_std_name in OPTIONAL_COLS.items():
        if opt_key in actual_cols_lower:
            mapping[actual_cols_lower[opt_key]] = opt_std_name

    if missing:
        raise ValueError(f"Chybí povinné sloupce: {', '.join(missing)}")
        
    return df.rename(columns=mapping)

def process_file(input_path, output_path):
    print(f"--- Zpracovávám soubor: {input_path} ---")
    
    try:
        # 1. Autodetekce listu (z utils)
        sheet_name = find_best_sheet(input_path)
        df = pd.read_excel(input_path, sheet_name=sheet_name)
        
        # 2. Normalizace
        df = normalize_columns(df)
        
        # 3. Filtr Storage location
        allowed_locations = ["F010", "F070"]
        df = df[df["Storage location"].astype(str).isin(allowed_locations)].copy()
        
        # 4. Smazání sloupců
        cols_to_drop = ["Storage location"]
        if "Plant" in df.columns:
            cols_to_drop.append("Plant")
        df = df.drop(columns=cols_to_drop, errors='ignore')
        
        # 5. Uspořádání
        ordered_cols = ["Material", "Material description", "Batch", "Total Quantity"]
        remaining_cols = [c for c in df.columns if c not in ordered_cols]
        df = df[ordered_cols + remaining_cols]
        
        # 6. Přejmenování
        rename_map = {"Material description": "Nazev", "Total Quantity": "Mnozstvi_SAP"}
        df = df.rename(columns=rename_map)
        
        # 7. Sort
        df["Mnozstvi_SAP"] = pd.to_numeric(df["Mnozstvi_SAP"], errors='coerce').fillna(0)
        df = df.sort_values(by="Mnozstvi_SAP", ascending=False, kind='mergesort')
        
        # 8. Smazat první datový řádek po sortu
        if len(df) > 0:
            df = df.iloc[1:]
        
        # 9. Ořezat na 4 sloupce
        df = df.iloc[:, :4]
        
        # --- EXPORT DAT (Pandas) ---
        # Použijeme openpyxl engine pro kompatibilitu
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='SAP', index=False)
            
        # --- FORMÁTOVÁNÍ TABULKY (z utils) ---
        create_excel_table(output_path, 'SAP', 'tbl_SAP')
        
        print(f"✅ Hotovo. Uloženo do: {output_path}")

    except Exception as e:
        print(f"❌ Chyba při zpracování: {e}")
        sys.exit(1)

if __name__ == "__main__":
    input_dir = "sklady_porovnani/input"
    output_dir = "sklady_porovnani/output"
    filename = "SAP.xlsx"
    
    infile = os.path.join(input_dir, filename)
    outfile = os.path.join(output_dir, filename)
    
    os.makedirs(output_dir, exist_ok=True)

    if not os.path.exists(infile):
        print(f"❌ CHYBA: Soubor '{infile}' neexistuje.")
        sys.exit(1)

    process_file(infile, outfile)