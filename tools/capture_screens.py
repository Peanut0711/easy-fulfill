"""화면 캡처와 전후 비교 도구(UI 개편 0단계, docs/ui-redesign-plan.md).

사용:
    .venv\\Scripts\\python.exe tools\\capture_screens.py
        → captures/<커밋>/ 에 배율 100·125·150% 캡처와 index.html
    .venv\\Scripts\\python.exe tools\\capture_screens.py --scales 100 --out captures/before
    .venv\\Scripts\\python.exe tools\\capture_screens.py --compare captures/before captures/after
        → captures/compare-before-after/index.html (전·후·차이 이미지), 바뀐 화면 목록 출력
    ... --compare A B --fail-on-diff   # 한 장이라도 다르면 종료 코드 1(구조 분리 단계의 「모양 그대로」 검사)

- 창은 화면에 띄우지 않는다(``WA_DontShowOnScreen``). 배율은 위젯을 지정한 장치 픽셀 비율로 그려
  흉내 낸다. 창 크기는 FHD 작업 영역(배율별)에서 앱이 정하는 기본 크기다(``shared.window_geometry``).
- 결과가 PC마다 달라지지 않도록 PC 상태를 고정한다: 임시 폴더를 작업 폴더로 쓰고(이 PC의
  database/·output/을 읽거나 쓰지 않음), Google 연결은 막고 상태는 「연결됨」으로, 앱 버전은
  ``0.0.0``으로, 문서 날짜는 2026-01-05로, 쿠팡 로그인 정보는 「판매자」로, 배송추적 목록은 빈 목록으로
  고정한다. 구조를 옮겨 아래 ``_worker``가 고치는 메서드 이름이 바뀌면 이 도구도 함께 고친다.
- 값은 Qt 논리 픽셀이다. 실제 모니터 배율과는 글꼴 힌팅이 조금 다를 수 있다(docs/known-risks.md).
"""

import argparse
import html
import json
import os
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SCALES = (100, 125, 150)
# FHD(1920×1080) 화면의 배율별 작업 영역(작업 표시줄 제외, 논리 픽셀). docs/gui-window-settings-plan.md 「크기 기준」
FHD_WORK_AREAS = {100: (1920, 1032), 125: (1536, 826), 150: (1280, 688)}
GROUPS = ("tabs", "settings", "states", "menus", "dialogs")
FIXED_VERSION = "0.0.0"
FIXED_DOCUMENT_DATE = (2026, 1, 5)
WORKER_TIMEOUT_S = 300


def parse_scales(text):
    """``"100,125"`` → ``(100, 125)``. 알 수 없는 배율이면 ValueError."""
    scales = tuple(int(part) for part in str(text).split(",") if part.strip())
    unknown = [scale for scale in scales if scale not in FHD_WORK_AREAS]
    if not scales or unknown:
        raise ValueError(f"배율은 {sorted(FHD_WORK_AREAS)} 중에서 고른다: {text}")
    return scales


def capture_window_size(scale, minimum):
    """배율 ``scale``의 FHD 작업 영역에서 앱이 정하는 기본 창 크기(논리 픽셀)."""
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    from shared.window_geometry import default_window_size

    return default_window_size(FHD_WORK_AREAS[scale], minimum)


def safe_name(text):
    keep = [ch if ch.isalnum() or ch in "-_" else "_" for ch in str(text)]
    return "".join(keep).strip("_") or "unnamed"


# ---------------------------------------------------------------- 비교(Pillow, Qt 불필요)

def compare_images(before_path, after_path, diff_path=None):
    """두 PNG를 비교해 ``{"status", "changed_pixels", "bbox"}``를 돌려준다.

    ``status``: ``"same"``·``"changed"``·``"size"``(크기가 다름). ``diff_path``가 있고 다르면
    바뀐 화소를 빨갛게 칠한 이미지(나중 이미지 위)를 저장한다.
    """
    from PIL import Image, ImageChops

    with Image.open(before_path) as a_img, Image.open(after_path) as b_img:
        a = a_img.convert("RGB")
        b = b_img.convert("RGB")
    if a.size != b.size:
        return {"status": "size", "changed_pixels": None, "bbox": None,
                "before_size": list(a.size), "after_size": list(b.size)}
    mask = ImageChops.difference(a, b).convert("L").point(lambda v: 255 if v else 0)
    bbox = mask.getbbox()
    if bbox is None:
        return {"status": "same", "changed_pixels": 0, "bbox": None}
    changed = mask.histogram()[255]
    if diff_path is not None:
        faded = Image.blend(b, Image.new("RGB", b.size, "white"), 0.6)
        faded.paste(Image.new("RGB", b.size, (220, 0, 0)), mask=mask)
        faded.save(diff_path)
    return {"status": "changed", "changed_pixels": changed, "bbox": list(bbox)}


def _png_set(directory):
    directory = Path(directory)
    return {path.relative_to(directory).as_posix() for path in directory.rglob("*.png")}


def compare_dirs(before_dir, after_dir, out_dir=None):
    """두 캡처 폴더를 같은 상대 경로끼리 비교한다. 한쪽에만 있으면 ``"added"``·``"removed"``."""
    before_dir, after_dir = Path(before_dir), Path(after_dir)
    before, after = _png_set(before_dir), _png_set(after_dir)
    rows = []
    for rel in sorted(before | after):
        if rel not in after:
            rows.append({"path": rel, "status": "removed"})
            continue
        if rel not in before:
            rows.append({"path": rel, "status": "added"})
            continue
        diff_path = None
        if out_dir is not None:
            diff_path = Path(out_dir) / "diff" / rel
            diff_path.parent.mkdir(parents=True, exist_ok=True)
        row = compare_images(before_dir / rel, after_dir / rel, diff_path)
        row["path"] = rel
        rows.append(row)
    return rows


def write_compare_html(rows, before_dir, after_dir, out_dir):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    def rel_url(path):
        return Path(os.path.relpath(path, out_dir)).as_posix()

    order = {"changed": 0, "size": 0, "added": 1, "removed": 1, "same": 2}
    parts = [
        "<!doctype html><meta charset='utf-8'><title>캡처 비교</title>",
        "<style>body{font-family:'Malgun Gothic',sans-serif;margin:16px}"
        "td{vertical-align:top;padding:4px}img{max-width:420px;border:1px solid #ccc}"
        ".same{color:#888}.changed,.size{color:#c0392b;font-weight:600}</style>",
        f"<h1>캡처 비교</h1><p>전: {html.escape(str(before_dir))}<br>후: {html.escape(str(after_dir))}</p>",
        "<table><tr><th>화면</th><th>전</th><th>후</th><th>차이</th></tr>",
    ]
    for row in sorted(rows, key=lambda r: (order.get(r["status"], 3), r["path"])):
        rel = row["path"]
        status = row["status"]
        info = status
        if status == "changed":
            info = f"changed · {row['changed_pixels']} px · {row['bbox']}"
        elif status == "size":
            info = f"size {row['before_size']} → {row['after_size']}"
        cells = []
        for base in (Path(before_dir), Path(after_dir), out_dir / "diff"):
            path = base / rel
            cells.append(f"<td><img src='{html.escape(rel_url(path))}'></td>" if path.exists() else "<td>—</td>")
        parts.append(f"<tr><td class='{status}'>{html.escape(rel)}<br>{html.escape(info)}</td>{''.join(cells)}</tr>")
    parts.append("</table>")
    (out_dir / "index.html").write_text("\n".join(parts), encoding="utf-8")


# ---------------------------------------------------------------- 캡처 결과 묶음

def write_capture_html(out_dir, manifest):
    out_dir = Path(out_dir)
    parts = [
        "<!doctype html><meta charset='utf-8'><title>화면 캡처</title>",
        "<style>body{font-family:'Malgun Gothic',sans-serif;margin:16px}"
        "figure{display:inline-block;margin:8px;vertical-align:top}"
        "img{max-width:460px;border:1px solid #ccc}figcaption{font-size:12px}</style>",
        f"<h1>화면 캡처 · {html.escape(manifest.get('label', ''))}</h1>",
        f"<p>커밋 {html.escape(manifest.get('commit', ''))} · 앱 {html.escape(manifest.get('app_version', ''))}"
        f" · {html.escape(manifest.get('captured_at', ''))}</p>",
    ]
    for scale_info in manifest.get("scales", []):
        scale = scale_info["scale"]
        parts.append(f"<h2>{scale}% · 창 {scale_info.get('window_size')}</h2>")
        if scale_info.get("errors"):
            parts.append("<p style='color:#c0392b'>" + "<br>".join(map(html.escape, scale_info["errors"])) + "</p>")
        for shot in scale_info.get("shots", []):
            src = html.escape(f"{scale}/{shot['file']}")
            parts.append(f"<figure><a href='{src}'><img src='{src}'></a>"
                         f"<figcaption>{html.escape(shot['file'])} · {shot['size']}</figcaption></figure>")
    (out_dir / "index.html").write_text("\n".join(parts), encoding="utf-8")


def _git(*args):
    try:
        return subprocess.run(["git", *args], cwd=ROOT, capture_output=True, text=True,
                              encoding="utf-8", check=True).stdout.strip()
    except (OSError, subprocess.CalledProcessError):
        return ""


def _commit_label():
    commit = _git("rev-parse", "--short", "HEAD") or "unknown"
    dirty = bool(_git("status", "--porcelain", "--untracked-files=no"))
    return commit + ("-dirty" if dirty else "")


def _real_app_version():
    text = (ROOT / "sheets" / "version_gate.py").read_text(encoding="utf-8")
    for line in text.splitlines():
        if line.startswith("APP_VERSION"):
            return line.split("=", 1)[1].strip().strip("\"'")
    return ""


def run_capture(scales, out_dir, platform="windows", groups=GROUPS):
    """배율마다 자식 프로세스로 캡처하고 manifest·index.html을 쓴다. manifest dict를 돌려준다."""
    out_dir = Path(out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    manifest = {
        "label": out_dir.name,
        "commit": _commit_label(),
        "branch": _git("rev-parse", "--abbrev-ref", "HEAD"),
        "app_version": _real_app_version(),
        "captured_at": datetime.now().isoformat(timespec="seconds"),
        "platform": platform,
        "scales": [],
    }
    env = dict(os.environ)
    env.pop("QT_SCALE_FACTOR", None)
    for scale in scales:
        scale_dir = out_dir / str(scale)
        if scale_dir.exists():
            shutil.rmtree(scale_dir)
        command = [sys.executable, "-I", str(Path(__file__).resolve()), "--worker",
                   "--scale", str(scale), "--out", str(scale_dir), "--platform", platform,
                   "--groups", ",".join(groups)]
        try:
            proc = subprocess.run(command, capture_output=True, text=True, encoding="utf-8",
                                  errors="replace", timeout=WORKER_TIMEOUT_S, env=env)
            output = proc.stdout
        except subprocess.TimeoutExpired as exc:
            output = exc.stdout or ""
            if isinstance(output, bytes):
                output = output.decode("utf-8", "replace")
        result = None
        for line in output.splitlines():
            if line.startswith("RESULT "):
                result = json.loads(line[len("RESULT "):])
        if result is None:
            tail = (output or "")[-2000:]
            result = {"scale": scale, "shots": [], "errors": [f"캡처 프로세스 실패: {tail}"]}
        manifest["scales"].append(result)
        print(f"{scale}%: {len(result['shots'])}장, 오류 {len(result.get('errors', []))}건", flush=True)
        for error in result.get("errors", []):
            print(f"  ! {error.splitlines()[0] if error else error}", flush=True)
    (out_dir / "capture.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    write_capture_html(out_dir, manifest)
    return manifest


# ---------------------------------------------------------------- 자식 프로세스(Qt)

def _worker(scale, out_dir, platform, groups):
    """한 배율의 화면을 모두 캡처하고 ``RESULT {json}`` 한 줄을 출력한다."""
    os.environ["QT_QPA_PLATFORM"] = platform
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")  # 부모가 RESULT 줄을 UTF-8로 읽는다
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    out_dir = Path(out_dir).resolve()  # 아래에서 작업 폴더를 임시 폴더로 바꾸므로 먼저 절대 경로로
    out_dir.mkdir(parents=True, exist_ok=True)
    workdir = Path(tempfile.mkdtemp(prefix="ef-capture-"))
    os.chdir(workdir)

    import importlib.util
    from unittest.mock import patch

    from PySide6.QtCore import QDate, Qt
    from PySide6.QtGui import QPixmap
    from PySide6.QtWidgets import QApplication, QMessageBox, QWidget

    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication.instance() or QApplication([])
    # main()과 같은 앱 전체 설정을 쓴다. 1단계에서 테마 적용 함수가 바뀌면 여기도 함께 바꾼다.
    from shared.ui_style import apply_app_style

    apply_app_style(app)

    dpr = scale / 100.0
    shots, errors = [], []

    def render(widget, name):
        try:
            widget.ensurePolished()
            size = widget.size()
            pixmap = QPixmap(max(1, round(size.width() * dpr)), max(1, round(size.height() * dpr)))
            pixmap.setDevicePixelRatio(dpr)
            pixmap.fill(Qt.GlobalColor.white)
            widget.render(pixmap)
            file_name = f"{name}.png"
            if not pixmap.save(str(out_dir / file_name)):
                raise OSError(f"저장 실패: {out_dir / file_name}")
            shots.append({"file": file_name, "size": [size.width(), size.height()]})
        except Exception as exc:  # 한 화면이 실패해도 나머지는 찍는다
            errors.append(f"{name}: {type(exc).__name__}: {exc}")

    def settle(rounds=6):
        for _ in range(rounds):
            app.processEvents()

    def wait_background(timeout_s=10.0, quiet_s=1.0):
        """시작 때 띄운 백그라운드 작업(막힌 네트워크로 곧 실패)이 끝나 화면 문구가 정해질 때까지 기다린다.

        타이머로 늦게 시작하는 작업도 있어, 작업 스레드가 ``quiet_s`` 동안 하나도 돌지 않아야 끝난 것으로 본다.
        """
        import time

        from PySide6.QtCore import QThread, QThreadPool

        def busy():
            if QThreadPool.globalInstance().activeThreadCount() > 0:
                return True
            owners = [app, *window_holder]
            return any(thread.isRunning() for owner in owners for thread in owner.findChildren(QThread))

        deadline = time.monotonic() + timeout_s
        quiet_since = time.monotonic()
        while time.monotonic() < deadline:
            settle()
            now = time.monotonic()
            if busy():
                quiet_since = now
            elif now - quiet_since >= quiet_s:
                break
            time.sleep(0.05)
        settle()

    window_holder = []

    def show_hidden(widget):
        widget.setAttribute(Qt.WidgetAttribute.WA_DontShowOnScreen, True)
        widget.show()
        settle()

    class _FixedDate(QDate):
        @staticmethod
        def currentDate():
            return QDate(*FIXED_DOCUMENT_DATE)

    import sheets.google_sheets_oauth as oauth
    import sheets.version_gate as version_gate

    silent = lambda *a, **k: QMessageBox.StandardButton.Ok  # noqa: E731
    spec = importlib.util.spec_from_file_location("easy_fulfill_capture", ROOT / "easy-fulfill.py")
    module = importlib.util.module_from_spec(spec)
    patches = [
        patch.object(oauth, "get_authorized_gspread_client", side_effect=RuntimeError("캡처: 네트워크 금지")),
        patch.object(oauth, "get_oauth_status", return_value=("ok", "연결됨", "캡처용 고정 상태")),
        patch.object(version_gate, "APP_VERSION", FIXED_VERSION),
        patch.object(QMessageBox, "warning", side_effect=silent),
        patch.object(QMessageBox, "critical", side_effect=silent),
        patch.object(QMessageBox, "information", side_effect=silent),
        patch.object(QMessageBox, "question", side_effect=lambda *a, **k: QMessageBox.StandardButton.No),
    ]
    try:
        import documents.quotation_statement_widget as quotation_widget

        patches.append(patch.object(quotation_widget, "QDate", _FixedDate))
    except Exception as exc:
        errors.append(f"문서 날짜 고정 실패: {exc}")
    for item in patches:
        item.start()

    window_size = None
    try:
        spec.loader.exec_module(module)
        main_window_cls = module.MainWindow
        main_window_cls._detail_session_label = lambda self: "저장된 로그인 정보: 판매자 (작업 시 권한 확인)"
        if hasattr(main_window_cls, "_quick_output_dir"):
            main_window_cls._quick_output_dir = lambda self: workdir / "output"
        if hasattr(main_window_cls, "_load_tracking_list"):
            # 막힌 네트워크로 실패하는 목록 읽기와 탭 진입 때의 건수 갱신이 경쟁해 건수 문구가 매번 달라진다.
            # 읽기를 하지 않아 빈 목록(처음 실행한 PC와 같은 상태)으로 고정한다.
            main_window_cls._load_tracking_list = lambda self: None
        try:  # 컨트롤러가 직접 부르는 경로(탭 진입·필터·주기 갱신)도 막는다
            import tracking.controller as tracking_controller

            tracking_controller.TrackingController._load_tracking_list = lambda self: None
        except Exception as exc:
            errors.append(f"배송추적 목록 고정 실패: {exc}")

        window = main_window_cls()
        window_holder.append(window)
        show_hidden(window)
        wait_background()
        tabs = window.ui.tabWidget
        for index in range(tabs.count()):  # 첫 진입 때 만드는 화면까지 만든 뒤 크기를 정한다
            tabs.setCurrentIndex(index)
            settle()
        tabs.setCurrentIndex(0)
        settle()
        minimum = window.minimumSizeHint()
        window_size = capture_window_size(scale, (minimum.width(), minimum.height()))
        window.resize(*window_size)
        wait_background()

        if "tabs" in groups:
            for index in range(tabs.count()):
                tabs.setCurrentIndex(index)
                settle()
                render(window, f"tab-{index + 1}-{safe_name(tabs.widget(index).objectName())}")

        stack = getattr(window.ui, "stackedWidget_settings", None)
        if "settings" in groups and stack is not None:
            tabs.setCurrentWidget(window.ui.tab_settings)
            nav = getattr(window.ui, "listWidget_settings_nav", None)
            for index in range(stack.count()):
                if nav is not None and index < nav.count():
                    nav.setCurrentRow(index)
                stack.setCurrentIndex(index)
                settle()
                render(window, f"settings-{index + 1}-{safe_name(stack.widget(index).objectName())}")

        if "states" in groups:
            tabs.setCurrentIndex(0)
            settle()
            for kind, text in (("ok", "우체국 접수 엑셀을 만들었습니다."),
                               ("warn", "송장 3건이 주문과 매칭되지 않았습니다."),
                               ("error", "공유 시트에 저장하지 못했습니다. 네트워크를 확인하세요.")):
                try:
                    window.notify(kind, text, actions=(("폴더 열기", lambda: None),))
                    settle()
                    render(window, f"state-notice-{kind}")
                except Exception as exc:
                    errors.append(f"state-notice-{kind}: {type(exc).__name__}: {exc}")
            notice = getattr(window, "notice_bar", None) or getattr(window, "_notice_bar", None)
            if notice is not None and hasattr(notice, "hide"):
                notice.hide()
            status_label = getattr(window, "status_label", None)
            if status_label is not None:  # 다음 화면에 마지막 알림 문구가 남지 않게
                status_label.setText("준비")
            try:
                window._show_busy_processing_overlay("주문을 불러오는 중입니다", "잠시만 기다려 주세요")
                settle()
                render(window, "state-busy-overlay")
                window._hide_busy_processing_overlay()
            except Exception as exc:
                errors.append(f"state-busy-overlay: {type(exc).__name__}: {exc}")
            # 시작 오버레이(_startup_overlay)는 만들기만 하고 앱에서 보이지 않아 찍지 않는다(계획 4-7에서 정리).

        if "menus" in groups:
            for index, action in enumerate(window.menuBar().actions()):
                menu = action.menu()
                if menu is None:
                    continue
                menu.ensurePolished()
                menu.adjustSize()
                render(menu, f"menu-{index + 1}-{safe_name(action.text().replace('&', ''))}")

        if "dialogs" in groups:
            def dialog(name, factory):
                try:
                    widget = factory()
                    show_hidden(widget)
                    render(widget, f"dialog-{name}")
                    widget.close()
                    widget.deleteLater()
                    settle()
                except Exception as exc:
                    errors.append(f"dialog-{name}: {type(exc).__name__}: {exc}")

            def message_box(icon, title, text, buttons):
                box = QMessageBox(icon, title, text, buttons, window)
                return box

            dialog("question", lambda: message_box(
                QMessageBox.Icon.Question, "확인", "우체국에 실제로 접수할까요?\n접수 후에는 되돌릴 수 없습니다.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No))
            dialog("warning", lambda: message_box(
                QMessageBox.Icon.Warning, "경고", "먼저 표에서 행을 선택해 주세요.", QMessageBox.StandardButton.Ok))
            dialog("kpost-address-search", lambda: module.KpostAddressSearchDialog(window))

            def unavailable(*_args, **_kwargs):
                raise RuntimeError("캡처: 네트워크 금지")

            try:
                from ui.gift_event_settings import GiftEventSettingsDialog
                from ui.gift_history import GiftHistoryDialog

                dialog("gift-event-settings", lambda: GiftEventSettingsDialog(unavailable, window))
                dialog("gift-history", lambda: GiftHistoryDialog(unavailable, window))
            except Exception as exc:
                errors.append(f"사은품 창: {type(exc).__name__}: {exc}")

            sample = workdir / "sample-detail.html"
            sample.write_text("<h2>상품 상세</h2><p>캡처용 예시 문단입니다.</p><hr><p>두 번째 문단</p>", encoding="utf-8")
            dialog("detail-html-editor", lambda: module.DetailHtmlEditorDialog(sample, window))
    except Exception as exc:
        errors.append(f"메인 창: {type(exc).__name__}: {exc}")
    finally:
        print("RESULT " + json.dumps({
            "scale": scale, "dpr": dpr, "window_size": list(window_size) if window_size else None,
            "font": [app.font().family(), app.font().pointSizeF()],
            "shots": shots, "errors": errors,
        }, ensure_ascii=False), flush=True)
        sys.stdout.flush()
        os.chdir(ROOT)
        shutil.rmtree(workdir, ignore_errors=True)
        # 백그라운드 QThread 정리 단계를 건너뛴다(easy-fulfill.py main()과 같은 이유)
        os._exit(0)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--scales", default=",".join(map(str, DEFAULT_SCALES)))
    parser.add_argument("--out", help="결과 폴더(기본 captures/<커밋>, 비교면 captures/compare-<전>-<후>)")
    parser.add_argument("--platform", default="windows")
    parser.add_argument("--groups", default=",".join(GROUPS), help=f"찍을 묶음: {','.join(GROUPS)}")
    parser.add_argument("--compare", nargs=2, metavar=("BEFORE", "AFTER"))
    parser.add_argument("--fail-on-diff", action="store_true", help="비교에서 하나라도 다르면 종료 코드 1")
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--scale", type=int, help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    groups = tuple(part for part in args.groups.split(",") if part)
    unknown_groups = [group for group in groups if group not in GROUPS]
    if unknown_groups:
        parser.error(f"알 수 없는 묶음: {unknown_groups}")

    if args.worker:
        _worker(args.scale, args.out, args.platform, groups)
        return 0

    if args.compare:
        before, after = (Path(path) for path in args.compare)
        out_dir = Path(args.out) if args.out else ROOT / "captures" / f"compare-{before.name}-{after.name}"
        rows = compare_dirs(before, after, out_dir)
        write_compare_html(rows, before, after, out_dir)
        different = [row for row in rows if row["status"] != "same"]
        print(f"같음 {len(rows) - len(different)}장 · 다름 {len(different)}장 → {out_dir / 'index.html'}")
        for row in different:
            print(f"  {row['status']:8s} {row['path']}")
        return 1 if (args.fail_on_diff and different) else 0

    try:
        scales = parse_scales(args.scales)
    except ValueError as exc:
        parser.error(str(exc))
    out_dir = Path(args.out) if args.out else ROOT / "captures" / _commit_label()
    manifest = run_capture(scales, out_dir, args.platform, groups)
    print(f"→ {out_dir / 'index.html'}")
    return 1 if any(info.get("errors") for info in manifest["scales"]) else 0


if __name__ == "__main__":
    sys.exit(main())
