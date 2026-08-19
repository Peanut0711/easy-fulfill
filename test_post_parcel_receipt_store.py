"""우체국 실제 접수 이력 저장·출력 후보 검증."""

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from post_parcel import ParcelReceipt
from post_parcel_receipt_store import ParcelReceiptStore, ReceiptStoreError


RECEIPT = ParcelReceipt(
    order_no="EFREAL-TEST-001",
    req_no="REQ-001",
    res_no="RES-001",
    regi_no="REGI-001",
    rechecked=True,
)


class ParcelReceiptStoreTests(unittest.TestCase):
    def test_real_receipt_is_saved_as_pending_print_candidate(self):
        with TemporaryDirectory() as directory:
            store = ParcelReceiptStore(Path(directory) / "receipts.sqlite3")

            saved = store.record_real_receipt(RECEIPT)

            self.assertEqual(saved.print_status, "PENDING")
            self.assertEqual([candidate.regi_no for candidate in store.list_pending_prints()], ["REGI-001"])

    def test_same_receipt_is_idempotent_but_different_receipt_same_order_is_blocked(self):
        with TemporaryDirectory() as directory:
            store = ParcelReceiptStore(Path(directory) / "receipts.sqlite3")
            store.record_real_receipt(RECEIPT)

            same = store.record_real_receipt(RECEIPT)
            self.assertEqual(same.regi_no, RECEIPT.regi_no)

            with self.assertRaises(ReceiptStoreError):
                store.record_real_receipt(
                    ParcelReceipt(
                        order_no=RECEIPT.order_no,
                        req_no="REQ-002",
                        res_no="RES-002",
                        regi_no="REGI-002",
                        rechecked=True,
                    ),
                )


if __name__ == "__main__":
    unittest.main()
