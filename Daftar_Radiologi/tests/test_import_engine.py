import unittest
import os
import sys
import shutil
import tempfile
import openpyxl

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
APP_ROOT = os.path.dirname(CURRENT_DIR)
if APP_ROOT not in sys.path:
    sys.path.insert(0, APP_ROOT)

from core.import_engine import parse_excel_file, auto_suggest_mapping, process_data_migration

class TestImportEngine(unittest.TestCase):
    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.sample_xlsx = os.path.join(self.test_dir, "sample_legacy.xlsx")
        
        # Cipta fail Excel sampel
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Rekod 2026"
        ws.append(["Tarikh", "No Xray", "Nama Pesakit", "No IC", "Umur", "Jantina", "Pemeriksaan", "Bahagian"])
        ws.append(["2026-07-15", "5001", "AHMAD BIN ALI", "900101015555", "36", "L", "GENERAL RADIOGRAPHY", "CHEST PA"])
        ws.append(["2026-07-16", "5002", "SITI KHALIFAH", "950202026666", "31", "P", "ULTRASOUND", "ABDOMEN FULL"])
        wb.save(self.sample_xlsx)
        wb.close()

    def tearDown(self):
        if os.path.exists(self.test_dir):
            shutil.rmtree(self.test_dir)

    def test_parse_excel_file(self):
        res = parse_excel_file(self.sample_xlsx)
        self.assertTrue(res.get("success"))
        self.assertIn("Tarikh", res.get("headers"))
        self.assertIn("Nama Pesakit", res.get("headers"))

    def test_auto_suggest_mapping(self):
        headers = ["Tarikh", "No Xray", "Nama Pesakit", "No IC", "Pemeriksaan", "Bahagian"]
        suggested = auto_suggest_mapping(headers)
        self.assertEqual(suggested.get("tarikh"), 0)
        self.assertEqual(suggested.get("no_xray"), 1)
        self.assertEqual(suggested.get("nama"), 2)
        self.assertEqual(suggested.get("no_ic"), 3)

    def test_process_data_migration(self):
        headers = ["Tarikh", "No Xray", "Nama Pesakit", "No IC", "Umur", "Jantina", "Pemeriksaan", "Bahagian"]
        mapping = {
            "tarikh": 0,
            "no_xray": 1,
            "nama": 2,
            "no_ic": 3,
            "umur": 4,
            "jantina": 5,
            "jenis_pemeriksaan": 6,
            "bahagian_pemeriksaan": 7
        }
        res = process_data_migration(self.sample_xlsx, mapping, sheet_name="Rekod 2026")
        self.assertTrue(res.get("success"))
        self.assertEqual(res.get("success_count"), 2)

if __name__ == "__main__":
    unittest.main()
