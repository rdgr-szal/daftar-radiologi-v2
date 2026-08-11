import os
import sys
import time
import datetime
import unittest

# Tambah path root aplikasi
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
APP_ROOT = os.path.dirname(CURRENT_DIR)
sys.path.insert(0, APP_ROOT)

import pydicom
from pydicom.dataset import Dataset
from pydicom.sequence import Sequence
import pynetdicom
from pynetdicom import AE, sop_class

from core.dicom_engine import DicomMWLServerDaemon, MPPS_SOP_CLASS, MPPS_GENERAL_SOP_CLASS
from core.mpps_engine import (
    init_mpps_db,
    get_mpps_records_list,
    get_mpps_record_details,
    get_mpps_monthly_reject_summary,
    match_reject_category
)
from core.phris_engine import get_phris_matrix_data
from DaftarRadiologi import app

class TestMppsScpIntegration(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        init_mpps_db()
        cls.test_port = 11114
        cls.test_ae = "XRAY_TEST"
        cls.console_ae = "KK3SB_TEST"
        
        cls.daemon = DicomMWLServerDaemon.get_instance()
        cls.daemon.stop()
        time.sleep(0.5)
        ok, msg = cls.daemon.start(host="127.0.0.1", port=cls.test_port, ae_title=cls.test_ae)
        print(f"\n[Test Setup] Started DICOM SCP Server: {ok} ({msg})")
        time.sleep(0.5)

    @classmethod
    def tearDownClass(cls):
        cls.daemon.stop()
        print("\n[Test Teardown] Stopped DICOM SCP Server.")

    def test_01_reject_category_matcher(self):
        """Uji fungsi padanan 14 kategori penolakan rasmi PHRIS."""
        self.assertEqual(match_reject_category("Over exposure high kV"), "OVER EXPOSURE")
        self.assertEqual(match_reject_category("Imej terlalu gelap"), "OVER EXPOSURE")
        self.assertEqual(match_reject_category("Under exposed / pale image"), "UNDER EXPOSURE")
        self.assertEqual(match_reject_category("Patient motion / bergerak"), "PATIENT MOVEMENT")
        self.assertEqual(match_reject_category("Wrong positioning of hand"), "WRONG TECHNIQUE")
        self.assertEqual(match_reject_category("Wrong marker placed R instead of L"), "WRONG MARKER")
        self.assertEqual(match_reject_category("Collimation cut off anatomy"), "COLLIMATION ERROR")
        self.assertEqual(match_reject_category("Necklace artifact"), "PATIENT ARTIFACT")
        self.assertEqual(match_reject_category("DR Detector comm fault"), "DETECTOR FAULT")
        self.assertEqual(match_reject_category("Generator error code E04"), "EQUIPMENT FAULT")
        self.assertEqual(match_reject_category("Random unknown issue"), "MISCELLANEOUS")

    def test_02_mpps_n_create_and_n_set_completed(self):
        """Simulasi konsol Carestream menghantar N-CREATE diikuti N-SET COMPLETED."""
        ae_scu = AE(ae_title=self.console_ae.encode('ascii'))
        ae_scu.add_requested_context(MPPS_SOP_CLASS)
        ae_scu.add_requested_context(sop_class.ModalityPerformedProcedureStep)
        
        assoc = ae_scu.associate('127.0.0.1', self.test_port, ae_title=self.test_ae.encode('ascii'))
        self.assertTrue(assoc.is_established, "SCU Association failed to establish with MPPS SCP.")

        today_str = datetime.datetime.now().strftime("%Y%m%d")
        now_time_str = datetime.datetime.now().strftime("%H%M%S")
        sop_inst_uid = f"1.2.826.0.1.3680043.9.7128.999.{int(time.time())}.1"

        # 1. Hantar N-CREATE (IN PROGRESS)
        req_ds = Dataset()
        req_ds.PatientName = "RAMLI TEST"
        req_ds.PatientID = "901109105689"
        req_ds.AccessionNumber = "ACC-001"
        req_ds.Modality = "CR"
        req_ds.PerformedProcedureStepID = "PPS-001"
        req_ds.PerformedProcedureStepDescription = "CHEST PA"
        req_ds.PerformedStationAETitle = self.console_ae
        req_ds.PerformedStationName = "DRX_COMPASS"
        req_ds.PerformedProcedureStepStartDate = today_str
        req_ds.PerformedProcedureStepStartTime = now_time_str
        req_ds.PerformedProcedureStepStatus = "IN PROGRESS"

        status_create, rsp_ds = assoc.send_n_create(req_ds, MPPS_SOP_CLASS, sop_inst_uid)
        status_val = status_create.Status if hasattr(status_create, 'Status') else 0x0000
        self.assertEqual(status_val, 0x0000, "N-CREATE should return status 0x0000 (Success)")

        # Semak rekod tersimpan dalam SQLite
        rec = get_mpps_record_details(sop_inst_uid)
        self.assertIsNotNone(rec, "MPPS record should exist in database after N-CREATE")
        self.assertEqual(rec["status"], "IN PROGRESS")
        self.assertEqual(rec["patient_name"], "RAMLI TEST")
        self.assertEqual(rec["accession_number"], "ACC-001")

        # 2. Hantar N-SET (COMPLETED)
        set_ds = Dataset()
        set_ds.PerformedProcedureStepStatus = "COMPLETED"
        set_ds.PerformedProcedureStepEndDate = today_str
        set_ds.PerformedProcedureStepEndTime = datetime.datetime.now().strftime("%H%M%S")
        
        # Performed Series Sequence dengan 1 imej
        series_ds = Dataset()
        series_ds.SeriesInstanceUID = f"{sop_inst_uid}.1"
        series_ds.PerformingPhysicianName = "JURUXRAY 1"
        
        img_ds = Dataset()
        img_ds.ReferencedSOPClassUID = "1.2.840.10008.5.1.4.1.1.1"  # CR Image Storage
        img_ds.ReferencedSOPInstanceUID = f"{sop_inst_uid}.1.1"
        series_ds.ReferencedImageSequence = Sequence([img_ds])
        
        set_ds.PerformedSeriesSequence = Sequence([series_ds])

        status_set, rsp_set_ds = assoc.send_n_set(set_ds, MPPS_SOP_CLASS, sop_inst_uid)
        status_set_val = status_set.Status if hasattr(status_set, 'Status') else 0x0000
        self.assertEqual(status_set_val, 0x0000, "N-SET should return status 0x0000 (Success)")

        assoc.release()

        # Semak rekod dikemaskini
        rec_updated = get_mpps_record_details(sop_inst_uid)
        self.assertEqual(rec_updated["status"], "COMPLETED")
        self.assertEqual(rec_updated["total_images_count"], 1)

    def test_03_mpps_n_create_and_n_set_discontinued_with_reject(self):
        """Simulasi konsol Carestream menghantar N-SET DISCONTINUED dengan sebab penolakan imej."""
        ae_scu = AE(ae_title=self.console_ae.encode('ascii'))
        ae_scu.add_requested_context(MPPS_SOP_CLASS)
        
        assoc = ae_scu.associate('127.0.0.1', self.test_port, ae_title=self.test_ae.encode('ascii'))
        self.assertTrue(assoc.is_established)

        today_str = datetime.datetime.now().strftime("%Y%m%d")
        now_time_str = datetime.datetime.now().strftime("%H%M%S")
        sop_inst_uid = f"1.2.826.0.1.3680043.9.7128.999.{int(time.time())}.2"

        # 1. N-CREATE
        req_ds = Dataset()
        req_ds.PatientName = "SITI AISHAH"
        req_ds.PatientID = "880512105432"
        req_ds.AccessionNumber = "ACC-002"
        req_ds.Modality = "CR"
        req_ds.PerformedProcedureStepID = "PPS-002"
        req_ds.PerformedProcedureStepDescription = "PELVIS AP"
        req_ds.PerformedStationAETitle = self.console_ae
        req_ds.PerformedStationName = "DRX_COMPASS"
        req_ds.PerformedProcedureStepStartDate = today_str
        req_ds.PerformedProcedureStepStartTime = now_time_str
        req_ds.PerformedProcedureStepStatus = "IN PROGRESS"

        status_create, _ = assoc.send_n_create(req_ds, MPPS_SOP_CLASS, sop_inst_uid)
        self.assertEqual(status_create.Status if hasattr(status_create, 'Status') else 0, 0x0000)

        # 2. N-SET DISCONTINUED dengan reject reason
        set_ds = Dataset()
        set_ds.PerformedProcedureStepStatus = "DISCONTINUED"
        set_ds.PerformedProcedureStepEndDate = today_str
        set_ds.PerformedProcedureStepEndTime = datetime.datetime.now().strftime("%H%M%S")
        set_ds.CommentsOnThePerformedProcedureStep = "PATIENT MOVEMENT - Pesakit terbatuk semasa exposure dijalankan"

        status_set, _ = assoc.send_n_set(set_ds, MPPS_SOP_CLASS, sop_inst_uid)
        self.assertEqual(status_set.Status if hasattr(status_set, 'Status') else 0, 0x0000)

        assoc.release()

        # Semak rekod dan pendaftaran reject analysis
        rec = get_mpps_record_details(sop_inst_uid)
        self.assertEqual(rec["status"], "DISCONTINUED")
        self.assertTrue(len(rec.get("rejected_images", [])) > 0, "Rejected image log must be recorded")
        
        rej = rec["rejected_images"][0]
        self.assertEqual(rej["standard_category"], "PATIENT MOVEMENT", "Category should match PATIENT MOVEMENT")

    def test_04_phris_matrix_mpps_aggregation(self):
        """Uji pengagregatan data MPPS ke dalam matriks Reten PHRIS Seksyen 7."""
        current_year = datetime.datetime.now().year
        current_month_idx = datetime.datetime.now().month - 1

        matrix = get_phris_matrix_data(current_year)
        self.assertIn("penolakan", matrix)
        penolakan = matrix["penolakan"]

        # Pastikan 14 kategori wujud
        for cat in ["OVER EXPOSURE", "PATIENT MOVEMENT", "WRONG TECHNIQUE", "MISCELLANEOUS"]:
            self.assertIn(cat, penolakan)
            self.assertEqual(len(penolakan[cat]), 12)

        # Pastikan penolakan PATIENT MOVEMENT dikesan pada bulan semasa
        self.assertGreaterEqual(penolakan["PATIENT MOVEMENT"][current_month_idx], 1)
        self.assertGreaterEqual(penolakan["PENGULANGAN"][current_month_idx], 1)
        self.assertGreaterEqual(penolakan["JUMLAH IMEJ"][current_month_idx], 1)

    def test_05_flask_api_endpoints(self):
        """Uji endpoint REST API untuk MPPS records, audit JSON dan export CSV."""
        client = app.test_client()

        # 1. API Status
        res_stat = client.get('/api/dicom/status')
        self.assertEqual(res_stat.status_code, 200)
        d_stat = res_stat.get_json()
        self.assertTrue(d_stat["success"])
        self.assertTrue(d_stat["data"]["mpps_supported"])
        self.assertGreaterEqual(d_stat["data"]["mpps_records_count"], 1)

        # 2. API Records List
        res_list = client.get('/api/mpps/records?status=ALL')
        self.assertEqual(res_list.status_code, 200)
        d_list = res_list.get_json()
        self.assertTrue(d_list["success"])
        self.assertGreaterEqual(d_list["data"]["total"], 2)

        # 3. API Record Details
        records = d_list["data"]["records"]
        first_sop = records[0]["sop_instance_uid"]
        res_det = client.get(f'/api/mpps/records/{first_sop}')
        self.assertEqual(res_det.status_code, 200)
        d_det = res_det.get_json()
        self.assertTrue(d_det["success"])
        self.assertIn("raw_dataset", d_det["data"])

        # 4. API Export CSV
        res_exp = client.get('/api/mpps/export')
        self.assertEqual(res_exp.status_code, 200)
        self.assertEqual(res_exp.content_type, "text/csv; charset=utf-8")
        csv_text = res_exp.data.decode('utf-8')
        self.assertIn("Accession Number", csv_text)
        self.assertIn("ACC-001", csv_text)

if __name__ == '__main__':
    unittest.main()
