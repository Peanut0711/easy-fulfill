"""네이버 SmartEditor ONE 상세를 읽어 로컬 쿠팡용 미리보기를 만든다."""

from __future__ import annotations

import argparse
import hashlib
import html
import io
import json
import re
from collections import Counter
from html.parser import HTMLParser
from pathlib import Path

import requests
from PIL import Image

import naver_commerce
from google_sheets_oauth import get_authorized_gspread_client


SPREADSHEET_ID = "1F0l6FMjXvKXAR9WyDvxEWcRvji-TaJbBim_G12TJ2Pw"
CONFIG_SHEET_TITLE = "설정"
OUTPUT_ROOT = Path(__file__).resolve().parent / "output" / "detail-preview"
VOID_TAGS = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr"}
ZERO_WIDTH = re.compile("[\u200b\u200c\u200d\ufeff]")


class Node:
    def __init__(self, tag="", attrs=(), parent=None):
        self.tag = tag
        self.attrs = dict(attrs)
        self.parent = parent
        self.children = []


class TreeParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = Node("root")
        self.stack = [self.root]

    def handle_starttag(self, tag, attrs):
        node = Node(tag, attrs, self.stack[-1])
        self.stack[-1].children.append(node)
        if tag not in VOID_TAGS:
            self.stack.append(node)

    def handle_startendtag(self, tag, attrs):
        node = Node(tag, attrs, self.stack[-1])
        self.stack[-1].children.append(node)

    def handle_endtag(self, tag):
        for index in range(len(self.stack) - 1, 0, -1):
            if self.stack[index].tag == tag:
                self.stack = self.stack[:index]
                return

    def handle_data(self, data):
        if self.stack[-1].tag not in {"script", "style"}:
            self.stack[-1].children.append(data)


def classes(node):
    return set(node.attrs.get("class", "").split())


def walk(node):
    if isinstance(node, str):
        return
    yield node
    for child in node.children:
        yield from walk(child)


def clean_text(value):
    return re.sub(r"\s+", " ", ZERO_WIDTH.sub("", value)).strip()


def node_text(node):
    return clean_text("".join(child if isinstance(child, str) else node_text(child) for child in node.children))


def inline_html(node):
    parts = []
    for child in node.children:
        if isinstance(child, str):
            parts.append(html.escape(ZERO_WIDTH.sub("", child)))
        elif child.tag in {"b", "strong"}:
            parts.append(f"<strong>{inline_html(child)}</strong>")
        elif child.tag in {"i", "em"}:
            parts.append(f"<em>{inline_html(child)}</em>")
        elif child.tag == "br":
            parts.append("<br>")
        else:
            parts.append(inline_html(child))
    return "".join(parts)


def has_ancestor(node, tags):
    parent = node.parent
    while parent:
        if parent.tag in tags:
            return True
        parent = parent.parent
    return False


def render_text_component(component):
    large = any("se-fs-fs24" in classes(node) for node in walk(component))
    rendered = []
    for node in walk(component):
        if node.tag in {"ul", "ol"} and not has_ancestor(node, {"ul", "ol"}):
            items = [child for child in node.children if isinstance(child, Node) and child.tag == "li"]
            item_html = "".join(f"<li>{inline_html(item).strip()}</li>" for item in items if node_text(item))
            if item_html:
                rendered.append(f"<{node.tag}>{item_html}</{node.tag}>")
        elif node.tag == "p" and not has_ancestor(node, {"li"}) and node_text(node):
            tag = "h1" if large and not rendered else "p"
            rendered.append(f"<{tag}>{inline_html(node).strip()}</{tag}>")
    if not rendered and node_text(component):
        rendered.append(f"<p>{html.escape(node_text(component))}</p>")
    return "\n".join(rendered)


def render_table_component(component):
    table = next((node for node in walk(component) if node.tag == "table"), None)
    if not table:
        return ""
    rows = [node for node in walk(table) if node.tag == "tr"]
    rendered_rows = []
    for row_index, row in enumerate(rows):
        cells = [node for node in row.children if isinstance(node, Node) and node.tag in {"td", "th"}]
        if not cells:
            continue
        header = row_index == 0 and all(any(node.tag in {"b", "strong"} for node in walk(cell)) for cell in cells)
        tag = "th" if header else "td"
        rendered_cells = []
        for cell in cells:
            spans = "".join(f' {name}="{html.escape(cell.attrs[name], quote=True)}"' for name in ("colspan", "rowspan") if cell.attrs.get(name, "1") != "1")
            rendered_cells.append(f"<{tag}{spans}>{inline_html(cell).strip()}</{tag}>")
        rendered_rows.append(f"<tr>{''.join(rendered_cells)}</tr>")
    return f'<section class="table-block"><table><tbody>{"".join(rendered_rows)}</tbody></table></section>' if rendered_rows else ""


def render_section_title(component):
    title = node_text(component)
    return f"<section class=\"section-title\"><h2>{html.escape(title)}</h2></section>" if title else ""


def load_config():
    sheet = get_authorized_gspread_client().open_by_key(SPREADSHEET_ID).worksheet(CONFIG_SHEET_TITLE)
    return {
        row[0].strip(): (row[1] if len(row) > 1 else "").strip()
        for row in sheet.get_all_values()[1:]
        if row and row[0].strip()
    }


def fetch_product(product_no):
    config = load_config()
    token = naver_commerce.get_access_token(config["naver_client_id"], config["naver_client_secret"])
    response = requests.get(
        f"{naver_commerce.API_BASE}/v2/products/channel-products/{product_no}",
        headers={"Authorization": f"Bearer {token}"},
        timeout=naver_commerce.DEFAULT_TIMEOUT,
    )
    naver_commerce._raise_for_status_with_body(response)
    return response.json()


def download_image(url, path):
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    image = Image.open(io.BytesIO(response.content))
    image.load()
    path.write_bytes(response.content)
    return {"width": image.width, "height": image.height, "format": image.format, "bytes": len(response.content)}


def build_preview(product_no, product):
    origin = product.get("originProduct") or {}
    source = origin.get("detailContent") or ""
    if not source:
        raise RuntimeError("detailContent가 비어 있습니다.")

    parser = TreeParser()
    parser.feed(source)
    components = [node for node in walk(parser.root) if "se-component" in classes(node)]
    if not components:
        raise RuntimeError("SmartEditor ONE 컴포넌트를 찾지 못했습니다.")

    output_dir = OUTPUT_ROOT / str(product_no)
    image_dir = output_dir / "images"
    image_dir.mkdir(parents=True, exist_ok=True)
    body = []
    image_records = []
    image_index = 0

    for component in components:
        if "se-sectionTitle" in classes(component):
            rendered = render_section_title(component)
            if rendered:
                body.append(rendered)
            continue

        if "se-horizontalLine" in classes(component):
            body.append('<hr class="divider">')
            continue

        if "se-text" in classes(component):
            rendered = render_text_component(component)
            if rendered:
                body.append(f'<section class="text-block">{rendered}</section>')
            continue

        if "se-table" in classes(component):
            rendered = render_table_component(component)
            if rendered:
                body.append(rendered)
            continue

        image_nodes = [node for node in walk(component) if node.tag == "img" and node.attrs.get("src")]
        if not image_nodes:
            continue
        strip_classes = {name for node in walk(component) for name in classes(node) if "imageStrip-col-" in name}
        columns = 3 if "se-imageStrip-col-3" in strip_classes else 2 if "se-imageStrip-col-2" in strip_classes else 1
        rendered_images = []
        for image_node in image_nodes:
            image_index += 1
            path = image_dir / f"image-{image_index:02d}.jpg"
            metadata = download_image(image_node.attrs["src"], path)
            metadata.update({"index": image_index, "source": image_node.attrs["src"], "file": path.name})
            image_records.append(metadata)
            alt = html.escape(f"{origin.get('name', '상품 상세')} 이미지 {image_index}", quote=True)
            rendered_images.append(f'<img src="images/{path.name}" alt="{alt}">')
        body.append(f'<section class="image-block grid-{columns}">{"".join(rendered_images)}</section>')

    title = html.escape(origin.get("name") or f"네이버 상품 {product_no}")
    preview = f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>{title} - 쿠팡 상세 미리보기</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; background: #eef1f5; color: #222; font-family: Arial, 'Malgun Gothic', sans-serif; }}
    .notice {{ padding: 12px 16px; background: #fff4d6; color: #6b4c00; text-align: center; font-size: 14px; }}
    main {{ width: 100%; max-width: 780px; margin: 24px auto; padding: 42px 36px; background: white; }}
    .text-block {{ margin: 0 0 28px; font-size: 18px; line-height: 1.75; overflow-wrap: anywhere; }}
    h1 {{ margin: 0 0 28px; font-size: 26px; line-height: 1.4; text-align: center; }}
    h2 {{ margin: 0 0 20px; font-size: 24px; line-height: 1.4; text-align: center; }}
    p {{ margin: 0 0 12px; }}
    ul, ol {{ margin: 8px 0 16px; padding-left: 1.5em; }}
    li {{ margin: 5px 0; }}
    .image-block {{ display: grid; gap: 12px; margin: 0 0 30px; align-items: start; }}
    .grid-1 {{ grid-template-columns: minmax(0, 1fr); }}
    .grid-2 {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
    .grid-3 {{ grid-template-columns: repeat(3, minmax(0, 1fr)); }}
    img {{ display: block; width: auto; max-width: 100%; height: auto; margin: 0 auto; }}
    .table-block {{ margin: 0 0 30px; overflow-x: auto; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 16px; }}
    th, td {{ padding: 10px 12px; border: 1px solid #d6d6d6; text-align: left; vertical-align: top; }}
    th {{ background: #f5f5f5; font-weight: 700; }}
    .divider {{ border: 0; border-top: 1px solid #ddd; margin: 0 0 30px; }}
    @media (max-width: 560px) {{
      main {{ margin: 0; padding: 28px 18px; }}
      .text-block {{ font-size: 16px; line-height: 1.7; }}
      h1 {{ font-size: 22px; }}
      .grid-2, .grid-3 {{ grid-template-columns: minmax(0, 1fr); }}
    }}
  </style>
</head>
<body>
  <div class="notice">로컬 변환 미리보기 · 네이버/쿠팡 상품에는 반영되지 않았습니다.</div>
  <main>{''.join(body)}</main>
</body>
</html>
"""
    preview_path = output_dir / "coupang-preview.html"
    preview_path.write_text(preview, encoding="utf-8")

    report = {
        "channelProductNo": str(product_no),
        "name": origin.get("name", ""),
        "sourceSha256": hashlib.sha256(source.encode("utf-8")).hexdigest(),
        "sourceHtmlLength": len(source),
        "componentCount": len(components),
        "textComponentCount": sum("se-text" in classes(node) for node in components),
        "imageCount": len(image_records),
        "imageBytes": sum(record["bytes"] for record in image_records),
        "imageFormats": dict(Counter(record["format"] for record in image_records)),
        "tableComponentCount": sum("se-table" in classes(node) for node in components),
        "sectionTitleComponentCount": sum("se-sectionTitle" in classes(node) for node in components),
        "horizontalLineComponentCount": sum("se-horizontalLine" in classes(node) for node in components),
        "warnings": ["이미지 안의 글자는 선택 가능한 HTML 텍스트로 변환하지 않았습니다."],
        "images": image_records,
    }
    report_path = output_dir / "report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return preview_path, report_path, report


def self_test():
    parser = TreeParser()
    parser.feed('<div class="se-component se-text"><p><b>제목</b></p></div>')
    component = next(node for node in walk(parser.root) if "se-component" in classes(node))
    assert render_text_component(component) == "<p><strong>제목</strong></p>"
    parser.feed('<div class="se-component se-table"><table><tr><td><b>번호</b></td><td><b>설명</b></td></tr><tr><td>1</td><td>RS485</td></tr></table></div>')
    table = next(node for node in walk(parser.root) if "se-table" in classes(node))
    assert "<th><strong>번호</strong></th>" in render_table_component(table)
    parser.feed('<div class="se-component se-sectionTitle"><p>소제목</p></div>')
    section_title = next(node for node in walk(parser.root) if "se-sectionTitle" in classes(node))
    assert render_section_title(section_title) == '<section class="section-title"><h2>소제목</h2></section>'
    print("self-test: ok")


def main():
    arguments = argparse.ArgumentParser()
    arguments.add_argument("product_no", nargs="?")
    arguments.add_argument("--self-test", action="store_true")
    args = arguments.parse_args()
    if args.self_test:
        self_test()
        return
    if not args.product_no or not args.product_no.isdigit():
        arguments.error("숫자형 스마트스토어 상품번호가 필요합니다.")
    preview_path, report_path, report = build_preview(args.product_no, fetch_product(args.product_no))
    print(json.dumps({"preview": str(preview_path), "report": str(report_path), "summary": report}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
