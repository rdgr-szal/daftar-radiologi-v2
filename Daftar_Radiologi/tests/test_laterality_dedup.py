import os
import sys
import unittest

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
APP_ROOT = os.path.dirname(CURRENT_DIR)
sys.path.insert(0, APP_ROOT)

from core.dicom_engine import (
    dedup_laterality,
    build_procedure_name,
    build_procedure_id,
    add_to_dicom_worklist,
    load_dicom_worklist,
    clear_dicom_worklist
)

class TestLateralityDedup(unittest.TestCase):
    def test_dedup_laterality_basic(self):
        # Case from bug report
        self.assertEqual(dedup_laterality("RIGHT RADIUS ULNA RIGHT"), "RIGHT RADIUS ULNA")
        self.assertEqual(dedup_laterality("LEFT RADIUS ULNA LEFT"), "LEFT RADIUS ULNA")
        self.assertEqual(dedup_laterality("RIGHT KNEE RIGHT"), "RIGHT KNEE")
        self.assertEqual(dedup_laterality("CHEST"), "CHEST")
        self.assertEqual(dedup_laterality("CHEST PA"), "CHEST PA")
        self.assertEqual(dedup_laterality("RIGHT FOOT AP/LAT"), "RIGHT FOOT AP/LAT")
        self.assertEqual(dedup_laterality("KANAN FEMUR KANAN"), "KANAN FEMUR")

    def test_build_procedure_name(self):
        # 1. Body part without laterality + directional laterality
        self.assertEqual(build_procedure_name("RADIUS ULNA", "RIGHT"), "RIGHT RADIUS ULNA")
        self.assertEqual(build_procedure_name("RADIUS ULNA", "LEFT"), "LEFT RADIUS ULNA")
        self.assertEqual(build_procedure_name("FEMUR", "KANAN"), "RIGHT FEMUR")
        self.assertEqual(build_procedure_name("FEMUR", "KIRI"), "LEFT FEMUR")
        self.assertEqual(build_procedure_name("WRIST", "BILATERAL"), "BILATERAL WRIST")
        self.assertEqual(build_procedure_name("WRIST", "BOTH"), "BOTH WRIST")

        # 2. Body part already has laterality + redundant laterality param
        self.assertEqual(build_procedure_name("RIGHT RADIUS ULNA", "RIGHT"), "RIGHT RADIUS ULNA")
        self.assertEqual(build_procedure_name("LEFT RADIUS ULNA", "LEFT"), "LEFT RADIUS ULNA")
        self.assertEqual(build_procedure_name("RIGHT RADIUS ULNA", ""), "RIGHT RADIUS ULNA")
        self.assertEqual(build_procedure_name("RIGHT RADIUS ULNA RIGHT", "RIGHT"), "RIGHT RADIUS ULNA")

        # 3. Chest or non-directional exams
        self.assertEqual(build_procedure_name("CHEST", ""), "CHEST")
        self.assertEqual(build_procedure_name("CHEST", "PA"), "CHEST PA")
        self.assertEqual(build_procedure_name("ABDOMEN", "SUPINE"), "ABDOMEN SUPINE")

        # 4. Body part with laterality + projection/view
        self.assertEqual(build_procedure_name("RIGHT KNEE", "AP/LAT"), "RIGHT KNEE AP/LAT")
        self.assertEqual(build_procedure_name("LEFT ANKLE", "AP/LAT"), "LEFT ANKLE AP/LAT")

    def test_build_procedure_id(self):
        self.assertEqual(build_procedure_id("RIGHT RADIUS ULNA"), "RIGHTRADIUSULNA")
        self.assertEqual(build_procedure_id("LEFT KNEE"), "LEFTKNEE")
        self.assertEqual(build_procedure_id("CHEST PA"), "CHESTPA")
        self.assertEqual(build_procedure_id("CHEST"), "CHEST")

    def test_add_to_dicom_worklist_procedure_fields(self):
        clear_dicom_worklist()

        patient_data = {
            "nama": "RUN TEST MULTI EXAM",
            "ic_pasport": "880101-01-1234",
            "umur": 38,
            "jantina": "LELAKI",
            "nombor_xray": "0004",
            "bil_kes": "1",
            "bahagian_pemeriksaan": "RADIUS ULNA",
            "lateraliti": "RIGHT",
            "modality": "CR",
            "operator": "STAFF",
            "pegawai_rujukan": "DR X"
        }

        exam_records = [
            {
                "xray_no": "0004",
                "lateraliti": "RIGHT",
                "bahagian": "RIGHT RADIUS ULNA",
                "bil_kes": 1,
                "modality": "CR"
            }
        ]

        add_to_dicom_worklist(patient_data, exam_records)
        items = load_dicom_worklist()
        self.assertEqual(len(items), 1)

        item = items[0]
        # Verify no double "RIGHT"
        self.assertEqual(item["requested_procedure_id"], "RIGHTRADIUSULNA")
        self.assertEqual(item["requested_procedure_desc"], "RIGHT RADIUS ULNA")
        self.assertEqual(item["scheduled_step_desc"], "RIGHT RADIUS ULNA")

if __name__ == "__main__":
    unittest.main()
