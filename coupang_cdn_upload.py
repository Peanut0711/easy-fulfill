"""네이버 상세 미리보기의 이미지를 쿠팡 CDN에 올리고 HTML URL을 교체한다.

처음 실행하면 쿠팡 WING 로그인 창이 열립니다. 로그인 세션은 output/ 아래에만
저장되며, 이 도구는 상품 등록이나 임시저장을 호출하지 않습니다.
"""

from __future__ import annotations

import argparse
import ctypes
import json
import os
import re
import subprocess
import sys
import tempfile
import time
from ctypes import wintypes
from pathlib import Path
from urllib.parse import urlparse

from PIL import Image
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
OUTPUT_ROOT = ROOT / "output" / "detail-preview"
PROFILE_DIR = ROOT / "output" / "coupang-browser-profile"
AUTH_STATE_PATH = PROFILE_DIR / "easy-fulfill-auth-state.bin"
UPLOAD_URL = "https://wing.coupang.com/tenants/seller-web/file/resize/uploadV2"
WING_HOME = "https://wing.coupang.com/"
CDN_BASE = "https://image.coupangcdn.com/image/"
UPLOAD_RESPONSE_LOG_MAX_CHARS = 1_000


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


def _data_blob(data: bytes):
    buffer = (ctypes.c_byte * len(data)).from_buffer_copy(data)
    return _DataBlob(len(data), buffer), buffer


def _protect_for_current_windows_user(data: bytes):
    source, source_buffer = _data_blob(data)
    protected = _DataBlob()
    if not ctypes.windll.crypt32.CryptProtectData(
        ctypes.byref(source), None, None, None, None, 1, ctypes.byref(protected)
    ):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return ctypes.string_at(protected.pbData, protected.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(protected.pbData)


def _unprotect_for_current_windows_user(data: bytes):
    source, source_buffer = _data_blob(data)
    plain = _DataBlob()
    if not ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(source), None, None, None, None, 1, ctypes.byref(plain)
    ):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return ctypes.string_at(plain.pbData, plain.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(plain.pbData)


def save_coupang_session(context, page):
    """세션 쿠키를 현재 Windows 사용자만 읽을 수 있게 저장한다."""
    if os.name != "nt":
        raise RuntimeError("쿠팡 로그인 세션 저장은 Windows에서만 지원합니다.")
    state = context.storage_state()
    origin = urlparse(page.url)
    if origin.scheme == "https" and origin.hostname and origin.hostname.endswith("coupang.com"):
        state["easyFulfillSessionStorage"] = {
            f"{origin.scheme}://{origin.netloc}": page.evaluate("Object.fromEntries(Object.entries(sessionStorage))")
        }
    AUTH_STATE_PATH.write_bytes(_protect_for_current_windows_user(
        json.dumps(state, ensure_ascii=False).encode("utf-8")
    ))


def restore_coupang_session(context):
    """새 Chromium 실행 시 이전 로그인 쿠키를 먼저 되살린다."""
    if not AUTH_STATE_PATH.exists():
        return False
    try:
        state = json.loads(_unprotect_for_current_windows_user(AUTH_STATE_PATH.read_bytes()).decode("utf-8"))
        cookies = state.get("cookies", [])
        for cookie in cookies:
            if cookie.get("expires", -1) < 0:
                cookie.pop("expires", None)
        if cookies:
            context.add_cookies(cookies)
        for origin_state in state.get("origins", []):
            origin = origin_state.get("origin", "")
            local_storage = origin_state.get("localStorage", [])
            if local_storage:
                values = [(item["name"], item["value"]) for item in local_storage]
                context.add_init_script(
                    f"if (location.origin === {json.dumps(origin)}) for (const [key, value] of {json.dumps(values)}) localStorage.setItem(key, value);"
                )
        for origin, values in state.get("easyFulfillSessionStorage", {}).items():
            if values:
                context.add_init_script(
                    f"if (location.origin === {json.dumps(origin)}) for (const [key, value] of {json.dumps(list(values.items()))}) sessionStorage.setItem(key, value);"
                )
        return bool(cookies)
    except Exception:
        AUTH_STATE_PATH.unlink(missing_ok=True)
        print("저장된 쿠팡 로그인 세션을 복원하지 못했습니다. 새 로그인이 필요합니다.")
        return False


def load_report(product_no: str):
    output_dir = OUTPUT_ROOT / product_no
    report_path = output_dir / "report.json"
    if not report_path.exists():
        raise RuntimeError(f"미리보기 결과가 없습니다. 먼저 naver_detail_preview.py {product_no} 를 실행하세요.")
    report = json.loads(report_path.read_text(encoding="utf-8"))
    return output_dir, report


def _redact_upload_response_text(value: str):
    """오류 응답을 진단 로그에 남기되 인증값은 노출하지 않는다."""
    value = re.sub(
        r'(?i)(authorization|cookie|token|password|session)(["\']?\s*[:=]\s*["\']?)([^"\'\s;&<]+)',
        r"\1\2[REDACTED]",
        value,
    )
    value = re.sub(r"\s+", " ", value).strip()
    if len(value) > UPLOAD_RESPONSE_LOG_MAX_CHARS:
        return value[:UPLOAD_RESPONSE_LOG_MAX_CHARS] + " ...[truncated]"
    return value


def _read_upload_response(response):
    """JSON이 아닌 WING 오류 응답도 상태와 함께 운영자 로그에 남긴다."""
    content_type = response.headers.get("content-type", "(없음)")
    body = response.text()
    try:
        payload = json.loads(body)
    except json.JSONDecodeError as error:
        preview = _redact_upload_response_text(body) or "(빈 응답)"
        raise RuntimeError(
            "쿠팡 이미지 업로드 응답이 JSON이 아닙니다: "
            f"HTTP {response.status} {response.status_text}; "
            f"Content-Type={content_type}; URL={response.url}; 본문={preview}"
        ) from error
    if not isinstance(payload, dict):
        raise RuntimeError(
            "쿠팡 이미지 업로드 응답 형식이 객체가 아닙니다: "
            f"HTTP {response.status} {response.status_text}; "
            f"Content-Type={content_type}; 응답={_redact_upload_response_text(body)}"
        )
    return payload


def upload_images(context, output_dir: Path, report: dict):
    progress_path = output_dir / "coupang-cdn-progress.json"
    previous = []
    if progress_path.exists():
        previous = json.loads(progress_path.read_text(encoding="utf-8"))
    previous_by_source = {(item.get("source"), item.get("file")): item for item in previous}
    mapping = []
    for item in report["images"]:
        saved = previous_by_source.get((item.get("source"), item.get("file")))
        if saved and saved.get("cdnUrl"):
            mapping.append(saved)
            print(f"[{item['index']:02d}/{len(report['images']):02d}] 재사용 {saved['cdnUrl']}")
            continue
        image_path = output_dir / "images" / item["file"]
        upload_path = prepare_image_for_upload(image_path, output_dir / "upload-images", item["index"])
        response = context.request.post(
            UPLOAD_URL,
            headers={"Origin": "https://wing.coupang.com", "Referer": WING_HOME},
            multipart={
                "multipartFile": {
                    "name": image_path.name,
                    "mimeType": "image/jpeg",
                    "buffer": upload_path.read_bytes(),
                }
            },
            timeout=60_000,
        )
        payload = _read_upload_response(response)
        if not response.ok or not payload.get("success") or not payload.get("message"):
            raise RuntimeError(
                f"이미지 {item['index']} 업로드 실패: "
                f"HTTP {response.status} {response.status_text}; "
                f"Content-Type={response.headers.get('content-type', '(없음)')}; "
                f"응답={_redact_upload_response_text(json.dumps(payload, ensure_ascii=False))}"
            )
        mapping.append({
            **item,
            "cdnPath": payload["message"],
            "cdnUrl": CDN_BASE + payload["message"].lstrip("/"),
        })
        progress_path.write_text(json.dumps(mapping, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[{item['index']:02d}/{len(report['images']):02d}] {mapping[-1]['cdnUrl']}")
    return mapping


def prepare_image_for_upload(image_path: Path, upload_dir: Path, index: int):
    with Image.open(image_path) as opened:
        width, height = opened.size
        if 300 <= width <= 3_000 and height <= 30_000:
            return image_path
        scale = max(1, 300 / width)
        scale = min(scale, 3_000 / width, 30_000 / height)
        target = (round(width * scale), round(height * scale))
        if target[0] < 300 or target[1] > 30_000:
            raise RuntimeError(f"이미지 {index}의 비율은 쿠팡 허용 크기로 보정할 수 없습니다: {width}x{height}")
        image = opened.convert("RGBA")
        background = Image.new("RGBA", image.size, "white")
        background.alpha_composite(image)
        upload_dir.mkdir(parents=True, exist_ok=True)
        target_path = upload_dir / f"image-{index:02d}.jpg"
        background.convert("RGB").resize(target, Image.Resampling.LANCZOS).save(target_path, "JPEG", quality=95, optimize=True)
        print(f"이미지 {index} 보정: {width}x{height} → {target[0]}x{target[1]}")
        return target_path


def write_cdn_html(output_dir: Path, mapping: list[dict]):
    preview_path = output_dir / "coupang-preview.html"
    html = preview_path.read_text(encoding="utf-8")
    for item in mapping:
        html = html.replace(f'images/{item["file"]}', item["cdnUrl"])
    if re.search(r'src="images/', html):
        raise RuntimeError("일부 이미지 URL을 쿠팡 CDN URL로 교체하지 못했습니다.")
    html_path = output_dir / "coupang-cdn-preview.html"
    map_path = output_dir / "coupang-cdn-map.json"
    paste_path = output_dir / "coupang-paste.html"
    html_path.write_text(html, encoding="utf-8")
    map_path.write_text(json.dumps(mapping, ensure_ascii=False, indent=2), encoding="utf-8")
    paste_path.write_text(render_paste_html(html), encoding="utf-8")
    return html_path, map_path, paste_path


def render_paste_html(preview_html: str):
    match = re.search(r"<main>(.*)</main>", preview_html, re.DOTALL)
    if not match:
        raise RuntimeError("미리보기 본문을 찾지 못했습니다.")
    body = match.group(1)
    body = body.replace(
        '<section class="text-block">',
        '<div style="margin:0 0 28px;font-size:18px;line-height:1.75;overflow-wrap:anywhere">',
    )
    body = body.replace('<section class="section-title">', '<div style="margin:0 0 20px">')
    body = re.sub(
        r'<section class="image-block grid-[123]">',
        '<div style="margin:0 0 30px;text-align:center">',
        body,
    )
    body = body.replace('<section class="table-block">', '<div style="margin:0 0 30px;overflow-x:auto">')
    body = body.replace("</section>", "</div>")
    body = body.replace('<h1>', '<h1 style="margin:0 0 28px;font-size:26px;line-height:1.4;text-align:center">')
    body = body.replace('<h2>', '<h2 style="margin:0;font-size:24px;line-height:1.4;text-align:center">')
    body = body.replace('<hr class="divider">', '<hr style="border:0;border-top:1px solid #ddd;margin:0 0 30px">')
    body = re.sub(
        r"<img (?=src=)",
        '<img style="display:block;max-width:100%;height:auto;margin:0 auto 12px" ',
        body,
    )
    body = body.replace('<table>', '<table style="width:100%;border-collapse:collapse;font-size:16px">')
    body = re.sub(r"<(th|td)([^>]*)>", r'<\1 style="padding:10px 12px;border:1px solid #d6d6d6;text-align:left;vertical-align:top"\2>', body)
    body = body.replace('<th style="', '<th style="background:#f5f5f5;font-weight:700;')
    return f'<div style="max-width:780px;margin:0 auto;color:#222;font-family:Arial,Malgun Gothic,sans-serif">{body}</div>\n'


def wait_for_login(page):
    if _has_active_upload_session(page):
        print("쿠팡 WING 업로드 로그인 세션이 확인되었습니다.")
        return
    print("열린 Chrome 창에서 쿠팡 WING에 로그인해 주세요. 로그인되면 자동으로 업로드를 시작합니다.")
    if not _is_wing_url(page.url):
        page.goto(WING_HOME, wait_until="domcontentloaded", timeout=30_000)
    deadline = time.monotonic() + 300
    while True:
        if time.monotonic() > deadline:
            raise RuntimeError("쿠팡 WING 로그인 대기 시간이 초과되었습니다.")
        if _has_active_upload_session(page):
            break
        page.wait_for_timeout(1_000)
    print("쿠팡 WING 업로드 로그인 세션이 확인되었습니다.")


def _has_active_upload_session(page):
    """업로드 경로의 JSON 응답으로 인증을 확인하고, 로그인 리다이렉트는 거부한다."""
    try:
        # 이 URL은 파일 없이 GET하면 WING이 파일 오류 JSON을 돌려준다. 탭 이동으로
        # 확인하면 그 JSON이 사용자에게 노출되므로, 동일한 BrowserContext의 요청 API로
        # 최종 리다이렉트 URL만 확인한다. BrowserContext.request는 쿠키를 공유한다.
        response = page.context.request.get(UPLOAD_URL, timeout=30_000)
    except Exception as error:
        print(f"[로그인 확인] 업로드 권한 확인 요청 실패: {error}")
        return False
    if _is_active_upload_response(response):
        return True
    print(
        "[로그인 확인] 업로드 권한이 확인되지 않았습니다: "
        f"HTTP {response.status} {response.status_text}; "
        f"Content-Type={response.headers.get('content-type', '(없음)')}; URL={response.url}"
    )
    return False


def _is_active_upload_response(response):
    """파일 없는 업로드 GET의 JSON 오류 응답만 유효한 업로드 로그인으로 인정한다."""
    final_url = urlparse(response.url)
    upload_url = urlparse(UPLOAD_URL)
    content_type = response.headers.get("content-type", "").lower()
    return (
        response.status not in (401, 403)
        and final_url.hostname == upload_url.hostname
        and final_url.path == upload_url.path
        and "json" in content_type
    )


def _is_wing_url(url: str):
    return urlparse(url).hostname == "wing.coupang.com"


def launch_coupang_context(playwright, headless=False):
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    launch_options = {"headless": headless}
    if not headless:
        launch_options.update(args=["--window-size=1600,1000"], no_viewport=True)
    try:
        return playwright.chromium.launch_persistent_context(str(PROFILE_DIR), **launch_options)
    except Exception as error:
        if "Executable doesn't exist" not in str(error):
            raise
        print("쿠팡 WING용 Chromium을 처음 설치합니다. 잠시 기다려 주세요.")
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
        return playwright.chromium.launch_persistent_context(str(PROFILE_DIR), **launch_options)


def launch_coupang_upload_context(playwright):
    """저장 세션으로는 숨김 업로드하고, 로그인이 필요할 때만 창을 연다."""
    context = launch_coupang_context(playwright, headless=True)
    restore_coupang_session(context)
    page = context.pages[0] if context.pages else context.new_page()
    if _has_active_upload_session(page):
        print("저장된 쿠팡 로그인 세션으로 숨김 업로드를 진행합니다.")
        return context, page

    context.close()
    print("저장된 쿠팡 로그인 세션에 업로드 권한이 없습니다. WING 창에서 로그인해 주세요.")
    context = launch_coupang_context(playwright)
    restore_coupang_session(context)
    page = context.pages[0] if context.pages else context.new_page()
    wait_for_login(page)
    save_coupang_session(context, page)
    return context, page


def self_test():
    assert CDN_BASE + "vendor_inventory/test.jpg" == "https://image.coupangcdn.com/image/vendor_inventory/test.jpg"
    if os.name == "nt":
        assert _unprotect_for_current_windows_user(_protect_for_current_windows_user(b"session-test")) == b"session-test"
    paste_html = render_paste_html('<main><section class="image-block grid-2"><img src="a.jpg"></section></main>')
    assert "grid-2" not in paste_html and 'src="a.jpg"' in paste_html
    assert _redact_upload_response_text('token="secret"\nerror') == 'token="[REDACTED]" error'
    assert _is_wing_url(WING_HOME) and not _is_wing_url("https://xauth.coupang.com/login")
    class UploadResponse:
        def __init__(self, status, url, content_type):
            self.status = status
            self.url = url
            self.headers = {"content-type": content_type}

    assert _is_active_upload_response(UploadResponse(400, UPLOAD_URL, "application/json"))
    assert not _is_active_upload_response(UploadResponse(403, UPLOAD_URL, "text/html"))
    assert not _is_active_upload_response(UploadResponse(200, "https://wing.coupang.com/sso/login", "text/html"))
    with tempfile.TemporaryDirectory() as temp_dir:
        image_path = Path(temp_dir) / "small.png"
        Image.new("RGB", (120, 120), "white").save(image_path)
        resized = prepare_image_for_upload(image_path, Path(temp_dir) / "uploads", 1)
        assert Image.open(resized).size == (300, 300)
    print("self-test: ok")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("product_no", nargs="?")
    parser.add_argument("--upload", action="store_true", help="쿠팡 CDN에 실제 업로드합니다.")
    parser.add_argument("--render", action="store_true", help="기존 CDN 주소로 붙여넣기 HTML만 생성합니다.")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        self_test()
        return
    if not args.product_no or not args.product_no.isdigit():
        parser.error("숫자형 네이버 상품번호가 필요합니다.")
    if args.render:
        output_dir, _ = load_report(args.product_no)
        map_path = output_dir / "coupang-cdn-map.json"
        if not map_path.exists():
            raise RuntimeError("CDN 주소 매핑이 없습니다. 먼저 --upload를 실행하세요.")
        html_path, _, paste_path = write_cdn_html(output_dir, json.loads(map_path.read_text(encoding="utf-8")))
        print(json.dumps({"preview": str(html_path), "pasteHtml": str(paste_path)}, ensure_ascii=False))
        return
    if not args.upload:
        parser.error("실제 CDN 업로드는 --upload 옵션으로만 실행합니다.")

    output_dir, report = load_report(args.product_no)
    if not report.get("images"):
        raise RuntimeError("업로드할 이미지가 없습니다.")
    with sync_playwright() as playwright:
        context, page = launch_coupang_upload_context(playwright)
        try:
            mapping = upload_images(context, output_dir, report)
            html_path, map_path, paste_path = write_cdn_html(output_dir, mapping)
            print(json.dumps({"preview": str(html_path), "mapping": str(map_path), "pasteHtml": str(paste_path)}, ensure_ascii=False))
        finally:
            context.close()


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"오류: {error}", file=sys.stderr)
        raise SystemExit(1)
