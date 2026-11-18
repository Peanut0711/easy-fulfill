"""메인 창과 탭별 최소 크기를 잰다(창을 화면에 띄우지 않음).

사용:
    .venv\\Scripts\\python.exe tools\\measure_window_sizes.py          # 표 출력
    .venv\\Scripts\\python.exe tools\\measure_window_sizes.py --json   # JSON 한 줄

- 기본 플랫폼은 Windows(맑은 고딕)다. 화면 없는 환경은 ``--platform offscreen``으로
  재지만 글꼴이 달라 값이 커진다.
- 임시 폴더를 작업 폴더로 써서 이 PC의 database/ 파일을 읽거나 쓰지 않는다.
  Google 연결은 막는다.
- 값은 Qt 논리 픽셀이다. 크기 예산은 ``WINDOW_MIN_BUDGET``·``PAGE_MIN_BUDGET``.
"""

import argparse
import json
import os
import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

# docs/gui-window-settings-plan.md 「크기 기준」
WINDOW_MIN_BUDGET = (680, 560)
PAGE_MIN_BUDGET = (670, 440)
# 「설정」 탭 페이지 내용 최소 폭: 탭 예산 폭에서 분류 목록(140, 디자인 개편 4-6)·여백을 뺀 값. 넘으면 가로 스크롤이 생긴다.
SETTINGS_CONTENT_MAX_WIDTH = 510  # 창 최소 높이 = 탭 최소 높이 + 약 117(메뉴·탭바·상태바·여백)


def _size(widget_size):
    return [int(widget_size.width()), int(widget_size.height())]


def measure(platform="windows"):
    """창·탭·설정 페이지 최소 크기를 dict로 돌려준다(이 프로세스에서 QApplication을 만든다)."""
    import importlib.util
    from unittest.mock import patch

    os.environ["QT_QPA_PLATFORM"] = platform
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QApplication, QGroupBox, QLayout, QMessageBox, QWidget

    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication.instance() or QApplication([])
    from shared.ui_style import apply_app_style

    apply_app_style(app)  # main()과 같은 앱 스타일로 잰다
    import sheets.google_sheets_oauth as oauth

    spec = importlib.util.spec_from_file_location("easy_fulfill_measure", ROOT / "easy-fulfill.py")
    module = importlib.util.module_from_spec(spec)
    silent = lambda *a, **k: QMessageBox.StandardButton.Ok  # noqa: E731
    with patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("네트워크 금지")), \
            patch.object(QMessageBox, "warning", side_effect=silent), \
            patch.object(QMessageBox, "critical", side_effect=silent), \
            patch.object(QMessageBox, "information", side_effect=silent):
        spec.loader.exec_module(module)
        window = module.MainWindow()
        app.processEvents()
        tabs = window.ui.tabWidget
        result = {
            "platform": platform,
            "font": [app.font().family(), app.font().pointSizeF()],
            "window_size": _size(window.size()),
            "tabs": [],
            "settings_pages": [],
        }
        for index in range(tabs.count()):
            tabs.setCurrentIndex(index)  # 탭 첫 진입 때 만드는 화면(문서 만들기 등)까지 포함
            app.processEvents()
        # 숨겨진 탭은 자식이 바뀌어도 레이아웃 크기 캐시가 갱신되지 않는다(화면에 보일 때 갱신되어
        # 그때 창이 커진다). 위젯별 크기 캐시와 레이아웃 캐시를 모두 비워 실제로 보일 때의 값을 잰다.
        for widget in window.findChildren(QWidget):
            widget.updateGeometry()
        for layout in window.findChildren(QLayout):
            layout.invalidate()
        for index in range(tabs.count()):
            page = tabs.widget(index)
            result["tabs"].append({
                "name": page.objectName(),
                "title": tabs.tabText(index),
                "min": _size(page.minimumSizeHint()),
            })
        stack = getattr(window.ui, "stackedWidget_settings", None)
        if stack is not None:
            for index in range(stack.count()):
                page = stack.widget(index)
                content = page.widget() if hasattr(page, "widget") else None
                result["settings_pages"].append({
                    "name": page.objectName(),
                    "min": _size(page.minimumSizeHint()),
                    # 스크롤 없이 다 보이려면 필요한 크기(가로 스크롤은 없어야 한다)
                    "content": _size(content.minimumSizeHint()) if content is not None else None,
                    # 묶음 제목(실행 중 만드는 묶음 포함). 적용 범위 표시 검사에 쓴다.
                    "groups": [box.title() for box in page.findChildren(QGroupBox)],
                })
        tabs.setCurrentIndex(0)
        app.processEvents()
        result["window_min"] = _size(window.minimumSizeHint())
    return result


def _print_table(result):
    print(f"플랫폼 {result['platform']} · 글꼴 {result['font'][0]} {result['font'][1]:g}pt")
    print(f"창 크기 {result['window_size'][0]}×{result['window_size'][1]}  "
          f"창 최소 {result['window_min'][0]}×{result['window_min'][1]}  "
          f"(예산 {WINDOW_MIN_BUDGET[0]}×{WINDOW_MIN_BUDGET[1]})")
    print(f"탭 최소 크기 (예산 {PAGE_MIN_BUDGET[0]}×{PAGE_MIN_BUDGET[1]})")
    for row in result["tabs"]:
        width, height = row["min"]
        over = width > PAGE_MIN_BUDGET[0] or height > PAGE_MIN_BUDGET[1]
        print(f"  {row['name']:24s} {row['title']:10s} {width:4d}×{height:<4d}{'  ← 예산 초과' if over else ''}")
    if result["settings_pages"]:
        print(f"설정 페이지 내용 크기 (가로 예산 {SETTINGS_CONTENT_MAX_WIDTH}, 세로는 스크롤)")
        for row in result["settings_pages"]:
            width, height = row["content"] or row["min"]
            over = width > SETTINGS_CONTENT_MAX_WIDTH
            print(f"  {row['name']:36s} {width:4d}×{height:<4d}{'  ← 가로 초과' if over else ''}")


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--platform", default="windows")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(argv)
    directory = tempfile.mkdtemp(prefix="ef-measure-")
    os.chdir(directory)
    try:
        result = measure(args.platform)
        if args.json:
            print("RESULT " + json.dumps(result, ensure_ascii=False), flush=True)
        else:
            _print_table(result)
    finally:
        sys.stdout.flush()
        os.chdir(ROOT)
        shutil.rmtree(directory, ignore_errors=True)
        # 백그라운드 QThread 정리 단계를 건너뛴다(easy-fulfill.py main()과 같은 이유)
        os._exit(0)


if __name__ == "__main__":
    main()
