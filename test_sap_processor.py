import unittest
import pandas as pd
import openpyxl
import os
from sap_processor import process_file

class TestSapProcessor(unittest.TestCase):
    
    def setUp(self):
        self.input_file = "test_input_temp.xlsx"
        self.output_file = "test_output_temp.xlsx"
        
        data = {
            "mATERIAL": ["M1", "M2", "M3", "M4"],
            "Material description": ["Desc A", "Desc B", "Desc C", "Desc D"],
            "baTCH": ["B1", "B2", "B3", "B4"],
            "Total Quantity": ["10", "100", "5", "chyba"], 
            "Storage location": ["F010", "F070", "X999", "F010"], 
            "Plant": ["P1", "P1", "P1", "P1"],
            "NecoNavic": ["X", "X", "X", "X"]
        }
        
        df = pd.DataFrame(data)
        df.to_excel(self.input_file, sheet_name="Sheet1", index=False)
        print(f"\n[SETUP] Vytvořen testovací soubor: {self.input_file}")

    def test_process_logic(self):
        print("[TEST] Spouštím process_file...")
        process_file(self.input_file, self.output_file)
        
        self.assertTrue(os.path.exists(self.output_file), "Výstupní soubor nebyl vytvořen.")
        
        df_out = pd.read_excel(self.output_file, sheet_name="SAP")
        
        expected_cols = ["Material", "Název", "Batch", "Mnozstvi_SAP"]
        self.assertListEqual(list(df_out.columns), expected_cols, "Názvy nebo pořadí sloupců nesedí.")
        
        self.assertEqual(len(df_out), 2, "Počet řádků neodpovídá.")
        self.assertEqual(df_out.iloc[0]["Mnozstvi_SAP"], 10, "Řazení nebo smazání prvního řádku neproběhlo správně.")
        
        wb = openpyxl.load_workbook(self.output_file)
        ws = wb["SAP"]
        self.assertIn("tbl_SAP", ws.tables, "Tabulka 'tbl_SAP' nebyla nalezena.")
        
        print("[TEST] Všechny kontroly prošly úspěšně!")

    def tearDown(self):
        if os.path.exists(self.input_file):
            os.remove(self.input_file)
        if os.path.exists(self.output_file):
            os.remove(self.output_file)
        print("[TEARDOWN] Dočasné soubory smazány.")

if __name__ == "__main__":
    unittest.main()