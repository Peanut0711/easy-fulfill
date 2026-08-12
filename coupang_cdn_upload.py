"""네이버 상세 미리보기의 이미지를 쿠팡 CDN에 올리고 HTML URL을 교체한다.

처음 실행하면 쿠팡 WING 로그인 창이 열립니다. 로그인 세션은 output/ 아래에만
저장되며, 이 도구는 상품 등록이나 임시저장을 호출하지 않습니다.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import tempfile
import time
from pathlib import Path

from PIL import Image
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
OUTPUT_ROOT = ROOT / "output" / "detail-preview"
PROFILE_DIR = ROOT / "output" / "coupang-browser-profile"
UPLOAD_URL = "https://wing.coupang.com/tenants/seller-web/file/resize/uploadV2"
WING_HOME = "https://wing.coupang.com/"
CDN_BASE = "https://image.coupangcdn.com/image/"


def load_report(product_no: str):
    output_dir = OUTPUT_ROOT / product_no
    report_path = output_dir / "report.json"
    if not report_path.exists():
        raise RuntimeError(f"미리보기 결과가 없습니다. 먼저 naver_detail_preview.py {product_no} 를 실행하세요.")
    report = json.loads(report_path.read_text(encoding="utf-8"))
    return output_dir, report


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
        payload = response.json()
        if not response.ok or not payload.get("success") or not payload.get("message"):
            raise RuntimeError(f"이미지 {item['index']} 업로드 실패: {response.status} {payload}")
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
    body = re.sub(
        r'<section class="image-block grid-[123]">',
        '<div style="margin:0 0 30px;text-align:center">',
        body,
    )
    body = body.replace('<section class="table-block">', '<div style="margin:0 0 30px;overflow-x:auto">')
    body = body.replace("</section>", "</div>")
    body = body.replace('<h1>', '<h1 style="margin:0 0 28px;font-size:26px;line-height:1.4;text-align:center">')
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
    page.goto(WING_HOME)
    print("열린 Chrome 창에서 쿠팡 WING에 로그인해 주세요. 로그인되면 자동으로 업로드를 시작합니다.")
    page.wait_for_timeout(2_000)
    deadline = time.monotonic() + 300
    while "xauth.coupang.com" in page.url:
        if time.monotonic() > deadline:
            raise RuntimeError("쿠팡 WING 로그인 대기 시간이 초과되었습니다.")
        page.wait_for_timeout(1_000)
    if not re.match(r"https://wing\.coupang\.com/", page.url):
        raise RuntimeError(f"로그인 후 예상하지 못한 페이지로 이동했습니다: {page.url}")


def self_test():
    assert CDN_BASE + "vendor_inventory/test.jpg" == "https://image.coupangcdn.com/image/vendor_inventory/test.jpg"
    paste_html = render_paste_html('<main><section class="image-block grid-2"><img src="a.jpg"></section></main>')
    assert "grid-2" not in paste_html and 'src="a.jpg"' in paste_html
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
        PROFILE_DIR.mkdir(parents=True, exist_ok=True)
        context = playwright.chromium.launch_persistent_context(str(PROFILE_DIR), headless=False)
        page = context.pages[0] if context.pages else context.new_page()
        try:
            wait_for_login(page)
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
