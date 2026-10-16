"""별도 Python 프로세스로 사진을 업로드하며 Qt 화면을 계속 응답시킨다."""

import json
import sys
import tempfile
from pathlib import Path
from urllib.parse import urlsplit

from PySide6.QtCore import QProcess, QProcessEnvironment, Qt
from PySide6.QtWidgets import QDialog, QLabel, QProgressBar, QVBoxLayout


def validated_cdn_url(value):
    if not isinstance(value, str):
        raise ValueError("업로드 결과에 이미지 주소가 없습니다.")
    url = urlsplit(value)
    if (url.scheme != "https" or url.netloc != "image.coupangcdn.com"
            or not url.path.startswith("/image/") or url.path == "/image/"):
        raise ValueError("업로드 결과의 쿠팡 이미지 주소가 올바르지 않습니다.")
    return value


class ImageUploadDialog(QDialog):
    def __init__(self, parent, image_path, result_path):
        super().__init__(parent)
        self.setWindowTitle("사진 업로드")
        self.setWindowModality(Qt.WindowModal)
        self.setWindowFlags(Qt.Dialog | Qt.CustomizeWindowHint | Qt.WindowTitleHint)
        self.resize(520, 140)
        self.output = ""
        self.error = None
        layout = QVBoxLayout(self)
        label = QLabel("쿠팡 CDN에 사진을 업로드하고 있습니다.\n"
                       "로그인이 필요하면 열리는 WING 브라우저에서 로그인해 주세요.")
        label.setWordWrap(True)
        layout.addWidget(label)
        progress = QProgressBar()
        progress.setRange(0, 0)
        layout.addWidget(progress)
        self.process = QProcess(self)
        self.process.setWorkingDirectory(str(Path(__file__).resolve().parent))
        self.process.setProcessChannelMode(QProcess.MergedChannels)
        environment = QProcessEnvironment.systemEnvironment()
        environment.insert("PYTHONIOENCODING", "utf-8")
        environment.insert("PYTHONUNBUFFERED", "1")
        self.process.setProcessEnvironment(environment)
        self.process.readyReadStandardOutput.connect(self._read_output)
        self.process.finished.connect(self._finished)
        self.process.errorOccurred.connect(self._process_error)
        self.process.start(sys.executable, [str(Path(__file__).with_name("coupang_cdn_upload.py")),
                                           "--upload", "--local-image", str(image_path),
                                           "--result-json", str(result_path)])

    def _read_output(self):
        self.output = (self.output + bytes(self.process.readAllStandardOutput()).decode(
            "utf-8", errors="replace"))[-4000:]

    def _process_error(self, error):
        if error == QProcess.FailedToStart:
            self.error = "사진 업로드 프로그램을 시작하지 못했습니다: " + self.process.errorString()
            super().reject()

    def _finished(self, exit_code, status):
        self._read_output()
        if exit_code != 0 or status != QProcess.NormalExit:
            self.error = self.output.strip() or "사진 업로드 프로그램이 비정상 종료되었습니다."
            super().reject()
        else:
            self.accept()

    def reject(self):
        # Esc로 창만 닫고 진행 중인 로그인/업로드 프로세스를 버리지 않는다.
        if self.process.state() == QProcess.NotRunning:
            super().reject()

    def closeEvent(self, event):
        if self.process.state() != QProcess.NotRunning:
            event.ignore()
        else:
            super().closeEvent(event)


def upload_image_for_editor(parent, image_path):
    with tempfile.TemporaryDirectory(prefix="easy-fulfill-upload-") as directory:
        result_path = Path(directory) / "result.json"
        dialog = ImageUploadDialog(parent, image_path, result_path)
        try:
            if dialog.exec() != QDialog.Accepted:
                raise RuntimeError(dialog.error or "사진 업로드를 완료하지 못했습니다.")
            result = json.loads(result_path.read_text(encoding="utf-8"))
            return validated_cdn_url(result.get("cdnUrl"))
        finally:
            dialog.deleteLater()
