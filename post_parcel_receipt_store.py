"""우체국 실제 접수·운송장 출력 대상의 로컬 이력 저장소."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from contextlib import closing
import sqlite3

from post_parcel import ParcelReceipt


@dataclass(frozen=True)
class PrintCandidate:
    """운송장 포털 조회 전 안전하게 보여 줄 실제 접수 요약."""

    order_no: str
    req_no: str
    res_no: str
    regi_no: str
    received_at: str
    print_status: str


class ReceiptStoreError(RuntimeError):
    """실제 접수 이력의 저장·중복 검사 오류."""


def default_receipt_store_path() -> Path:
    """Git에 포함하지 않는 로컬 접수 이력 파일 경로."""
    return Path(__file__).resolve().parent / "output" / "post-parcel-receipts.sqlite3"


class ParcelReceiptStore:
    """실제 접수 결과와 아직 인쇄하지 않은 대상을 SQLite에 보관한다."""

    def __init__(self, db_path: Path | str | None = None):
        self.db_path = Path(db_path) if db_path else default_receipt_store_path()
        try:
            self.db_path.parent.mkdir(parents=True, exist_ok=True)
            self._initialize()
        except (OSError, sqlite3.Error) as error:
            raise ReceiptStoreError("실제 접수 이력 저장소를 열지 못했습니다.") from error

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.db_path)
        connection.row_factory = sqlite3.Row
        return connection

    def _initialize(self) -> None:
        with closing(self._connect()) as connection:
            with connection:
                connection.execute(
                    """
                    CREATE TABLE IF NOT EXISTS parcel_receipts (
                        order_no TEXT PRIMARY KEY,
                        req_no TEXT NOT NULL UNIQUE,
                        res_no TEXT NOT NULL,
                        regi_no TEXT NOT NULL UNIQUE,
                        received_at TEXT NOT NULL,
                        print_status TEXT NOT NULL DEFAULT 'PENDING',
                        print_command_at TEXT,
                        portal_printed_at TEXT
                    )
                    """,
                )

    def find_by_order_no(self, order_no: str) -> PrintCandidate | None:
        """같은 주문번호의 실제 접수 이력을 찾아 재접수 전 차단에 사용한다."""
        try:
            with closing(self._connect()) as connection:
                row = connection.execute(
                    """
                    SELECT order_no, req_no, res_no, regi_no, received_at, print_status
                    FROM parcel_receipts
                    WHERE order_no = ?
                    """,
                    (order_no,),
                ).fetchone()
        except sqlite3.Error as error:
            raise ReceiptStoreError("실제 접수 이력을 조회하지 못했습니다.") from error
        return self._row_to_candidate(row) if row else None

    def record_real_receipt(self, receipt: ParcelReceipt) -> PrintCandidate:
        """재조회가 끝난 실제 접수 1건을 인쇄 대기 상태로 저장한다."""
        existing = self.find_by_order_no(receipt.order_no)
        if existing:
            if existing.regi_no == receipt.regi_no:
                return existing
            raise ReceiptStoreError("같은 주문번호의 기존 실제 접수 이력이 있습니다.")

        received_at = datetime.now().isoformat(timespec="seconds")
        try:
            with closing(self._connect()) as connection:
                with connection:
                    connection.execute(
                        """
                        INSERT INTO parcel_receipts (
                            order_no, req_no, res_no, regi_no, received_at, print_status
                        ) VALUES (?, ?, ?, ?, ?, 'PENDING')
                        """,
                        (receipt.order_no, receipt.req_no, receipt.res_no, receipt.regi_no, received_at),
                    )
        except sqlite3.IntegrityError as error:
            raise ReceiptStoreError("동일한 실제 접수 이력이 이미 저장돼 있습니다.") from error
        except sqlite3.Error as error:
            raise ReceiptStoreError("실제 접수 이력을 저장하지 못했습니다.") from error

        return PrintCandidate(
            order_no=receipt.order_no,
            req_no=receipt.req_no,
            res_no=receipt.res_no,
            regi_no=receipt.regi_no,
            received_at=received_at,
            print_status="PENDING",
        )

    def list_pending_prints(self) -> list[PrintCandidate]:
        """아직 Windows 인쇄 명령을 보내지 않은 실제 접수 목록을 최근순으로 반환한다."""
        try:
            with closing(self._connect()) as connection:
                rows = connection.execute(
                    """
                    SELECT order_no, req_no, res_no, regi_no, received_at, print_status
                    FROM parcel_receipts
                    WHERE print_status = 'PENDING'
                    ORDER BY received_at DESC
                    """,
                ).fetchall()
        except sqlite3.Error as error:
            raise ReceiptStoreError("출력 대기 실제 접수 이력을 조회하지 못했습니다.") from error
        return [self._row_to_candidate(row) for row in rows]

    def list_portal_print_requests(self) -> list[PrintCandidate]:
        """포털 인쇄 요청은 했지만 포털 출력여부를 재확인하지 않은 이력을 반환한다."""
        try:
            with closing(self._connect()) as connection:
                rows = connection.execute(
                    """
                    SELECT order_no, req_no, res_no, regi_no, received_at, print_status
                    FROM parcel_receipts
                    WHERE print_status = 'PORTAL_PRINT_REQUESTED'
                    ORDER BY received_at DESC
                    """,
                ).fetchall()
        except sqlite3.Error as error:
            raise ReceiptStoreError("포털 인쇄 요청 이력을 조회하지 못했습니다.") from error
        return [self._row_to_candidate(row) for row in rows]

    def list_portal_print_confirmed(self) -> list[PrintCandidate]:
        """포털 출력여부를 이미 확인했지만 Windows 인쇄는 확인하지 않은 이력을 반환한다."""
        try:
            with closing(self._connect()) as connection:
                rows = connection.execute(
                    """
                    SELECT order_no, req_no, res_no, regi_no, received_at, print_status
                    FROM parcel_receipts
                    WHERE print_status = 'PORTAL_PRINT_CONFIRMED'
                    ORDER BY received_at DESC
                    """,
                ).fetchall()
        except sqlite3.Error as error:
            raise ReceiptStoreError("포털 출력 확인 이력을 조회하지 못했습니다.") from error
        return [self._row_to_candidate(row) for row in rows]

    def mark_portal_print_requested(self, regi_nos: list[str]) -> None:
        """포털 팝업의 인쇄 요청이 눌린 건을 재시도 대상에서 제외한다.

        우체국 포털은 이 시점에 이미 `신규출력`에서 제외할 수 있다. OZ Viewer
        확인이 실패해도 같은 건을 자동으로 다시 요청하지 않도록 즉시 기록한다.
        """
        normalized = [str(regi_no).strip() for regi_no in regi_nos if str(regi_no).strip()]
        if not normalized or len(set(normalized)) != len(normalized):
            raise ReceiptStoreError("포털 인쇄 요청으로 기록할 등기번호 목록이 올바르지 않습니다.")
        placeholders = ", ".join("?" for _ in normalized)
        try:
            with closing(self._connect()) as connection:
                with connection:
                    cursor = connection.execute(
                        f"""
                        UPDATE parcel_receipts
                        SET print_status = 'PORTAL_PRINT_REQUESTED', portal_printed_at = ?
                        WHERE regi_no IN ({placeholders}) AND print_status = 'PENDING'
                        """,
                        (datetime.now().isoformat(timespec="seconds"), *normalized),
                    )
                    if cursor.rowcount != len(normalized):
                        raise ReceiptStoreError(
                            "포털 인쇄 요청 이력을 안전하게 기록하지 못했습니다. 자동 재시도는 중단합니다.",
                        )
        except ReceiptStoreError:
            raise
        except sqlite3.Error as error:
            raise ReceiptStoreError("포털 인쇄 요청 이력을 저장하지 못했습니다.") from error

    def mark_portal_print_confirmed(self, regi_nos: list[str]) -> None:
        """포털 조회에서 출력여부가 출력으로 확인된 이력을 확정 상태로 바꾼다."""
        normalized = [str(regi_no).strip() for regi_no in regi_nos if str(regi_no).strip()]
        if not normalized or len(set(normalized)) != len(normalized):
            raise ReceiptStoreError("포털 출력 완료로 기록할 등기번호 목록이 올바르지 않습니다.")
        placeholders = ", ".join("?" for _ in normalized)
        try:
            with closing(self._connect()) as connection:
                with connection:
                    cursor = connection.execute(
                        f"""
                        UPDATE parcel_receipts
                        SET print_status = 'PORTAL_PRINT_CONFIRMED', portal_printed_at = ?
                        WHERE regi_no IN ({placeholders}) AND print_status = 'PORTAL_PRINT_REQUESTED'
                        """,
                        (datetime.now().isoformat(timespec="seconds"), *normalized),
                    )
                    if cursor.rowcount != len(normalized):
                        raise ReceiptStoreError("포털 출력 완료 이력을 안전하게 확정하지 못했습니다.")
        except ReceiptStoreError:
            raise
        except sqlite3.Error as error:
            raise ReceiptStoreError("포털 출력 완료 이력을 저장하지 못했습니다.") from error

    def mark_windows_print_requested(self, regi_nos: list[str]) -> None:
        """Windows 인쇄 창의 확인을 눌러 프린터 요청을 보낸 이력을 남긴다.

        스풀러·프린터의 물리 출력 완료는 이 이력만으로 증명하지 않는다. 같은 건의
        자동 재출력을 막기 위한 요청 시점 기록이다.
        """
        normalized = [str(regi_no).strip() for regi_no in regi_nos if str(regi_no).strip()]
        if not normalized or len(set(normalized)) != len(normalized):
            raise ReceiptStoreError("프린터 요청으로 기록할 등기번호 목록이 올바르지 않습니다.")
        placeholders = ", ".join("?" for _ in normalized)
        try:
            with closing(self._connect()) as connection:
                with connection:
                    cursor = connection.execute(
                        f"""
                        UPDATE parcel_receipts
                        SET print_status = 'WINDOWS_PRINT_REQUESTED', print_command_at = ?
                        WHERE regi_no IN ({placeholders}) AND print_status = 'PORTAL_PRINT_CONFIRMED'
                        """,
                        (datetime.now().isoformat(timespec="seconds"), *normalized),
                    )
                    if cursor.rowcount != len(normalized):
                        raise ReceiptStoreError("프린터 요청 이력을 안전하게 기록하지 못했습니다.")
        except ReceiptStoreError:
            raise
        except sqlite3.Error as error:
            raise ReceiptStoreError("프린터 요청 이력을 저장하지 못했습니다.") from error

    @staticmethod
    def _row_to_candidate(row: sqlite3.Row) -> PrintCandidate:
        return PrintCandidate(
            order_no=row["order_no"],
            req_no=row["req_no"],
            res_no=row["res_no"],
            regi_no=row["regi_no"],
            received_at=row["received_at"],
            print_status=row["print_status"],
        )
