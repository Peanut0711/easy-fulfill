"""네이버 상품번호 목록을 쿠팡 HTML 작성용 결과물로 일괄 변환한다.

쿠팡 CDN 이미지 업로드만 수행하며, 쿠팡 상품 등록·수정·임시저장 요청은 보내지 않는다.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

import coupang_cdn_upload
import naver_detail_preview


ROOT = Path(__file__).resolve().parent
OUTPUT_ROOT = ROOT / "output" / "detail-preview"


def product_numbers(values, list_path=None):
    if list_path:
        values = [*values, *Path(list_path).read_text(encoding="utf-8").replace(",", "\n").splitlines()]
    numbers = []
    for value in values:
        value = value.strip()
        if not value:
            continue
        if not value.isdigit():
            raise ValueError(f"숫자형 네이버 상품번호가 아닙니다: {value}")
        if value not in numbers:
            numbers.append(value)
    if not numbers:
        raise ValueError("네이버 상품번호를 하나 이상 입력하세요.")
    return numbers


def write_summary(results):
    path = OUTPUT_ROOT / f"batch-summary-{datetime.now():%Y%m%d-%H%M%S}.json"
    path.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def prepare(numbers):
    prepared, results = [], []
    for number in numbers:
        try:
            preview, report_path, report = naver_detail_preview.build_preview(number, naver_detail_preview.fetch_product(number))
            prepared.append((number, report))
            results.append({"productNo": number, "name": report["name"], "status": "prepared", "preview": str(preview), "report": str(report_path)})
            print(f"[준비] {number} · 이미지 {report['imageCount']}개")
        except Exception as error:
            results.append({"productNo": number, "status": "prepare_failed", "error": str(error)})
            print(f"[준비 실패] {number}: {error}", file=sys.stderr)
    return prepared, results


def upload(prepared, results):
    with sync_playwright() as playwright:
        context = coupang_cdn_upload.launch_coupang_context(playwright)
        coupang_cdn_upload.restore_coupang_session(context)
        page = context.pages[0] if context.pages else context.new_page()
        try:
            coupang_cdn_upload.wait_for_login(page)
            coupang_cdn_upload.save_coupang_session(context, page)
            for number, report in prepared:
                try:
                    output_dir = OUTPUT_ROOT / number
                    mapping = coupang_cdn_upload.upload_images(context, output_dir, report)
                    preview, mapping_path, paste_html = coupang_cdn_upload.write_cdn_html(output_dir, mapping)
                    result = next(item for item in results if item["productNo"] == number)
                    result.update({"status": "completed", "cdnPreview": str(preview), "cdnMapping": str(mapping_path), "pasteHtml": str(paste_html)})
                    print(f"[완료] {number} · {paste_html}")
                except Exception as error:
                    result = next(item for item in results if item["productNo"] == number)
                    result.update({"status": "upload_failed", "error": str(error)})
                    print(f"[업로드 실패] {number}: {error}", file=sys.stderr)
        finally:
            context.close()


def self_test():
    assert product_numbers(["12", "12", "34"]) == ["12", "34"]
    try:
        product_numbers(["bad"])
    except ValueError:
        print("self-test: ok")
        return
    raise AssertionError("잘못된 상품번호를 허용했습니다.")


def main():
    parser = argparse.ArgumentParser(description="네이버 상세를 쿠팡 HTML 작성용 결과물로 일괄 변환합니다.")
    parser.add_argument("product_no", nargs="*", help="네이버 스마트스토어 상품번호")
    parser.add_argument("--list", help="상품번호를 줄바꿈 또는 쉼표로 적은 UTF-8 텍스트 파일")
    parser.add_argument("--prepare-only", action="store_true", help="네이버 다운로드와 미리보기 생성까지만 수행합니다.")
    parser.add_argument("--upload", action="store_true", help="쿠팡 CDN 업로드까지 수행합니다.")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        self_test()
        return
    if args.prepare_only == args.upload:
        parser.error("--prepare-only 또는 --upload 중 하나를 지정하세요.")

    numbers = product_numbers(args.product_no, args.list)
    prepared, results = prepare(numbers)
    if args.upload and prepared:
        upload(prepared, results)
    summary = write_summary(results)
    print(json.dumps({"summary": str(summary), "completed": sum(item["status"] == "completed" for item in results)}, ensure_ascii=False))
    if any(item["status"].endswith("failed") for item in results):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
