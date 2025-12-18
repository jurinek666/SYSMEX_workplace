"""
Azure Functions App - SAP Excel Processor
Migrace z lokálního řešení na Azure Functions

Tento soubor obsahuje Azure Functions pro zpracování SAP Excel souborů.
"""

import azure.functions as func
import logging
import tempfile
import os
from io import BytesIO

# Import lokálních modulů (původní logika)
from utils import find_best_sheet, create_excel_table
from sap_processor import normalize_columns
import pandas as pd

# Inicializace Function App
app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

@app.route(route="sap_process", methods=["POST"])
def sap_process_http(req: func.HttpRequest) -> func.HttpResponse:
    """
    HTTP trigger pro zpracování SAP souboru.
    
    Očekává:
    - Multipart/form-data s Excel souborem (pole 'file')
    
    Vrací:
    - Zpracovaný Excel soubor jako attachment
    """
    logging.info('SAP Process HTTP trigger function zahájeno.')

    try:
        # Získání souboru z požadavku
        file = req.files.get('file')
        if not file:
            return func.HttpResponse(
                "Chyba: Soubor nebyl nalezen. Pošlete soubor v poli 'file'.",
                status_code=400
            )

        # Načtení obsahu do paměti
        file_content = file.read()
        
        # Zpracování
        output_bytes = process_sap_file_in_memory(file_content)
        
        # Vrácení zpracovaného souboru
        return func.HttpResponse(
            body=output_bytes,
            status_code=200,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={
                'Content-Disposition': 'attachment; filename=SAP_processed.xlsx'
            }
        )

    except Exception as e:
        logging.error(f"Chyba při zpracování: {str(e)}")
        return func.HttpResponse(
            f"Chyba při zpracování souboru: {str(e)}",
            status_code=500
        )


@app.blob_trigger(arg_name="myblob", 
                  path="input/{name}",
                  connection="AzureWebJobsStorage")
@app.blob_output(arg_name="outputblob",
                 path="output/{name}",
                 connection="AzureWebJobsStorage")
def sap_process_blob(myblob: func.InputStream, outputblob: func.Out[bytes]) -> None:
    """
    Blob trigger pro automatické zpracování při nahrání do Azure Blob Storage.
    
    Spouští se automaticky když je soubor nahrán do 'input' kontejneru.
    Výstup ukládá do 'output' kontejneru.
    """
    logging.info(f'Blob trigger: zpracovávám {myblob.name}, velikost: {myblob.length} bytes')

    try:
        # Načtení obsahu blobu
        file_content = myblob.read()
        
        # Zpracování
        output_bytes = process_sap_file_in_memory(file_content)
        
        # Uložení výstupu
        outputblob.set(output_bytes)
        
        logging.info(f'✅ Soubor {myblob.name} úspěšně zpracován')

    except Exception as e:
        logging.error(f'❌ Chyba při zpracování {myblob.name}: {str(e)}')
        raise


def process_sap_file_in_memory(file_content: bytes) -> bytes:
    """
    Core logika zpracování SAP souboru - pracuje pouze v paměti.
    
    Args:
        file_content: Obsah vstupního Excel souboru jako bytes
        
    Returns:
        Zpracovaný Excel soubor jako bytes
    """
    # Použití temporary file pro pandas/openpyxl operace
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_input:
        tmp_input.write(file_content)
        tmp_input_path = tmp_input.name

    try:
        # 1. Autodetekce listu
        sheet_name = find_best_sheet(tmp_input_path)
        df = pd.read_excel(tmp_input_path, sheet_name=sheet_name)
        
        # 2. Normalizace sloupců
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
        rename_map = {"Material description": "Název", "Total Quantity": "Mnozstvi_SAP"}
        df = df.rename(columns=rename_map)
        
        # 7. Sort
        df["Mnozstvi_SAP"] = pd.to_numeric(df["Mnozstvi_SAP"], errors='coerce').fillna(0)
        df = df.sort_values(by="Mnozstvi_SAP", ascending=False, kind='mergesort')
        
        # 8. Smazat první datový řádek po sortu
        if len(df) > 0:
            df = df.iloc[1:]
        
        # 9. Ořezat na 4 sloupce
        df = df.iloc[:, :4]
        
        # 10. Export do temporary file
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_output:
            tmp_output_path = tmp_output.name
        
        # Použít openpyxl engine pro kompatibilitu
        with pd.ExcelWriter(tmp_output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='SAP', index=False)
        
        # Formátování tabulky
        create_excel_table(tmp_output_path, 'SAP', 'tbl_SAP')
        
        # Načtení výsledku do paměti
        with open(tmp_output_path, 'rb') as f:
            output_bytes = f.read()
        
        return output_bytes

    finally:
        # Cleanup temporary files
        if os.path.exists(tmp_input_path):
            os.unlink(tmp_input_path)
        if 'tmp_output_path' in locals() and os.path.exists(tmp_output_path):
            os.unlink(tmp_output_path)
