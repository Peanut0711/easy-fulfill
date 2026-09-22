"""네이버 SmartEditor ONE 상세를 읽어 로컬 쿠팡용 미리보기를 만든다."""

from __future__ import annotations

import argparse
import hashlib
import html
import io
import json
import re
from urllib.parse import parse_qs, urlparse
from collections import Counter
from html.parser import HTMLParser
from pathlib import Path

import requests
from PIL import Image
from detail_image_layout import normalized_widths, render_image_group
from detail_text_style import portable_text_html, source_text_style, style_attribute

import naver_commerce
from google_sheets_oauth import get_authorized_gspread_client


SPREADSHEET_ID = "1F0l6FMjXvKXAR9WyDvxEWcRvji-TaJbBim_G12TJ2Pw"
CONFIG_SHEET_TITLE = "설정"
OUTPUT_ROOT = Path(__file__).resolve().parent / "output" / "detail-preview"
VOID_TAGS = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr"}
# ZWJ/ZWNJ는 이모지 조합과 문자 모양에 필요하므로 제거하지 않는다.
ZERO_WIDTH = re.compile("[\u200b\ufeff]")


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
    return "".join(render_text_node(child) for child in node.children)


def render_text_node(node):
    """문단/하위 목록의 경계를 유지하면서 SmartEditor 전용 래퍼를 벗긴다."""
    if isinstance(node, str):
        return html.escape(ZERO_WIDTH.sub("", node))
    if node.tag in {"script", "style"}:
        return ""
    if node.tag == "br":
        return "<br>"
    content = inline_html(node)
    tag = {"b": "strong", "i": "em", "strike": "s"}.get(node.tag, node.tag)
    style = source_text_style(node.attrs)
    extra = ""
    if tag == "a":
        href = node.attrs.get("href", "").strip()
        try:
            parsed_href = urlparse(href)
        except ValueError:
            return content
        if parsed_href.scheme not in {"http", "https"} or not parsed_href.netloc:
            return content
        extra = f' href="{html.escape(href, quote=True)}" target="_blank" rel="noopener noreferrer"'
    elif tag in {"ol", "li"}:
        key = "start" if tag == "ol" else "value"
        value = node.attrs.get(key, "")
        if re.fullmatch(r"-?\d+", value):
            extra += f' {key}="{value}"'
        if tag == "ol":
            if "reversed" in node.attrs:
                extra += " reversed"
            marker = {"1": "decimal", "a": "lower-alpha", "A": "upper-alpha", "i": "lower-roman", "I": "upper-roman"}.get(node.attrs.get("type"))
            if marker and "list-style-type" not in style:
                style += f";list-style-type:{marker}"
    if tag == "li":
        # SmartEditor는 들여쓰기만 있는 상위 li 아래에 실제 목록을 넣기도 한다.
        nested = any(isinstance(child, Node) and child.tag in {"ul", "ol"} for child in node.children)
        own_text = any(clean_text(child) if isinstance(child, str) else
                       (child.tag not in {"ul", "ol"} and node_text(child)) for child in node.children)
        if nested and not own_text:
            style += ";list-style-type:none;margin:0"
    if tag == "p" and not node_text(node) and not any(child.tag == "br" for child in walk(node)):
        content += "<br>"
    if tag == "span" and not style:
        return content
    if tag in {"p", "ul", "ol", "li", "strong", "em", "u", "s", "span", "a", "sub", "sup", "h1", "h2", "h3", "h4", "h5", "h6"}:
        return f"<{tag}{extra}{style_attribute(style.strip(';'))}>{content}</{tag}>"
    # 일반 div 안의 문단은 그대로 재귀 처리하고, div 자체에 의미 있는 서식이 있으면 보존한다.
    return f"<div{style_attribute(style)}>{content}</div>" if style else content


def render_text_component(component):
    rendered = render_text_node(component).strip()
    if rendered and not any(node.tag in {"p", "ul", "ol", "h1", "h2", "h3", "h4", "h5", "h6"} for node in walk(component)):
        rendered = f"<p>{rendered}</p>"
    return rendered


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


def youtube_video_id(url):
    """지원하는 YouTube 주소만 영상 ID로 정규화한다."""
    if not isinstance(url, str):
        return None
    try:
        parsed = urlparse(html.unescape(url).strip())
        if parsed.scheme not in {"http", "https", ""} or not parsed.netloc:
            return None
        if parsed.username or parsed.password or parsed.port not in {None, 80, 443}:
            return None
        host = (parsed.hostname or "").lower()
        parts = parsed.path.strip("/").split("/")
        if host in {"youtu.be", "www.youtu.be"} and len(parts) == 1:
            candidate = parts[0]
        elif host in {"youtube.com", "www.youtube.com", "m.youtube.com", "youtube-nocookie.com", "www.youtube-nocookie.com"}:
            if parts == ["watch"]:
                candidate = parse_qs(parsed.query).get("v", [""])[0]
            elif len(parts) == 2 and parts[0] in {"embed", "shorts", "live"}:
                candidate = parts[1]
            else:
                return None
        else:
            return None
    except ValueError:
        return None
    return candidate if re.fullmatch(r"[A-Za-z0-9_-]{11}", candidate) else None


def extract_youtube_videos(component):
    """SmartEditor의 숨겨진 oEmbed JSON과 직접 삽입한 iframe을 읽는다.

    원본 iframe/스크립트는 실행하거나 복사하지 않고 검증한 ID로 새로 만든다.
    data-module과 data-module-v2에 같은 영상이 있어도 한 번만 출력한다.
    """
    videos = []
    seen = set()

    def add(url, title=""):
        video_id = youtube_video_id(url)
        if video_id and video_id not in seen:
            seen.add(video_id)
            videos.append({"videoId": video_id, "url": f"https://www.youtube.com/watch?v={video_id}",
                           "embedUrl": f"https://www.youtube.com/embed/{video_id}",
                           "title": clean_text(title) if isinstance(title, str) and title.strip() else "제품 시연 영상"})

    def add_iframes(source, title=""):
        if not isinstance(source, str):
            return
        parser = TreeParser()
        parser.feed(source.replace('\\"', '"'))
        for node in walk(parser.root):
            if node.tag == "iframe":
                add(node.attrs.get("src"), title or node.attrs.get("title", ""))

    for node in walk(component):
        if node.tag == "iframe":
            add(node.attrs.get("src"), node.attrs.get("title", ""))
        for name in ("data-module-v2", "data-module"):
            raw = node.attrs.get(name)
            if not raw:
                continue
            try:
                module = json.loads(raw)
            except (ValueError, TypeError):
                continue
            if not isinstance(module, dict) or module.get("type") != "v2_oembed":
                continue
            data = module.get("data")
            if not isinstance(data, dict):
                continue
            title = data.get("title", "")
            add(data.get("inputUrl"), title)
            add_iframes(data.get("html"), title)
    return videos


def render_youtube_video(video):
    return (
        '<div class="video-block" style="margin:0 0 24px;text-align:center">'
        '<iframe width="680" height="383" style="display:block;width:100%;max-width:100%;'
        'height:auto;aspect-ratio:16 / 9;min-height:200px;border:0;margin:0 auto" '
        f'src="{html.escape(video["embedUrl"], quote=True)}" '
        f'title="{html.escape(video["title"], quote=True)}" '
        'allow="encrypted-media; fullscreen; picture-in-picture" '
        'referrerpolicy="strict-origin-when-cross-origin" allowfullscreen></iframe></div>'
    )


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


def image_rows(component):
    """SmartEditor의 이미지 행 경계와 열 개수를 보존한다."""
    row, row_owner, columns = [], None, 1
    for image in (node for node in walk(component) if node.tag == "img" and node.attrs.get("src")):
        owner, count = image, 1
        ancestor = image.parent
        while ancestor is not None:
            match = re.search(r'(?:^|\s)se-imageStrip(?:-col-)?([23])(?=\s|$)', ancestor.attrs.get("class", ""))
            if match:
                owner, count = ancestor, int(match[1])
                break
            if ancestor is component:
                break
            ancestor = ancestor.parent
        if row and (owner is not row_owner or len(row) == columns):
            yield row, columns
            row = []
        row_owner, columns = owner, count
        row.append(image)
    if row:
        yield row, columns


def image_width_percent(image, component):
    node = image
    while node is not None:
        if node is image or "se-module-image" in classes(node):
            match = re.search(r'(?:^|;)\s*width\s*:\s*(\d+(?:\.\d+)?)\s*%\s*(?:;|$)',
                              node.attrs.get("style", ""), re.IGNORECASE)
            if match and 0 < float(match[1]) <= 100:
                return float(match[1])
        if node is component:
            break
        node = node.parent
    return None


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
    image_groups = []
    video_records = []
    video_warnings = []
    skipped_video_count = 0
    image_index = 0

    for component_index, component in enumerate(components, 1):
        component_classes = classes(component)
        if component_classes & {"se-oembed", "se-video"} or any(node.tag == "iframe" for node in walk(component)):
            videos = extract_youtube_videos(component)
            if videos:
                for video in videos:
                    body.append(render_youtube_video(video))
                    video_records.append({**video, "componentIndex": component_index})
                continue
            skipped_video_count += 1
            video_warnings.append(f"구성요소 {component_index}: 지원하는 유튜브 영상 주소를 찾지 못해 재생기를 생성하지 않았습니다.")

        if "se-sectionTitle" in classes(component):
            rendered = render_section_title(component)
            if rendered:
                body.append(rendered)
            continue

        if "se-horizontalLine" in classes(component):
            body.append('<hr class="divider">')
            continue

        if "se-quotation" in classes(component):
            rendered = render_text_component(component)
            if rendered:
                body.append(f'<blockquote class="quote-block">{rendered}</blockquote>')
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

        for image_nodes, columns in image_rows(component):
            images = []
            for image_node in image_nodes:
                image_index += 1
                path = image_dir / f"image-{image_index:02d}.jpg"
                metadata = dict(download_image(image_node.attrs["src"], path))
                metadata.update({"index": image_index, "source": image_node.attrs["src"], "file": path.name})
                image_records.append(metadata)
                images.append({**metadata, "src": f"images/{path.name}",
                               "alt": f"{origin.get('name', '상품 상세')} 이미지 {image_index}"})
            if columns > 1:
                widths = normalized_widths(images, [image_width_percent(node, component) for node in image_nodes])
                body.append(render_image_group(images, widths))
                image_groups.append({"componentIndex": component_index,
                                     "imageIndices": [item["index"] for item in images], "widthPercentages": widths})
            else:
                image = images[0]
                body.append(f'<section class="image-block grid-1"><img src="{image["src"]}" '
                            f'alt="{html.escape(image["alt"], quote=True)}"></section>')

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
    .quote-block {{ margin: 0 0 28px; padding: 0 0 0 14px; border-left: 5px solid #555; font-size: 18px; line-height: 1.75; overflow-wrap: anywhere; }}
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
  <main>{portable_text_html(''.join(body))}</main>
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
        "quotationComponentCount": sum("se-quotation" in classes(node) for node in components),
        "imageCount": len(image_records),
        "imageGroups": image_groups,
        "videoCount": len(video_records),
        "skippedVideoComponentCount": skipped_video_count,
        "videos": video_records,
        "imageBytes": sum(record["bytes"] for record in image_records),
        "imageFormats": dict(Counter(record["format"] for record in image_records)),
        "tableComponentCount": sum("se-table" in classes(node) for node in components),
        "sectionTitleComponentCount": sum("se-sectionTitle" in classes(node) for node in components),
        "horizontalLineComponentCount": sum("se-horizontalLine" in classes(node) for node in components),
        "warnings": ["이미지 안의 글자는 선택 가능한 HTML 텍스트로 변환하지 않았습니다.", *video_warnings],
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
    parser = TreeParser()
    parser.feed('<div class="se-component se-text"><p><a href="https://example.com">제조사 문서</a></p></div>')
    component = next(node for node in walk(parser.root) if "se-component" in classes(node))
    assert 'href="https://example.com"' in render_text_component(component)
    parser.feed('<div class="se-component se-table"><table><tr><td><b>번호</b></td><td><b>설명</b></td></tr><tr><td>1</td><td>RS485</td></tr></table></div>')
    table = next(node for node in walk(parser.root) if "se-table" in classes(node))
    assert "<th><strong>번호</strong></th>" in render_table_component(table)
    parser.feed('<div class="se-component se-sectionTitle"><p>소제목</p></div>')
    section_title = next(node for node in walk(parser.root) if "se-sectionTitle" in classes(node))
    assert render_section_title(section_title) == '<section class="section-title"><h2>소제목</h2></section>'
    # 실제 누락 경로인 build_preview까지 실행해 인용구의 내용과 순서를 확인한다.
    import tempfile
    from unittest.mock import patch

    quote_lines = ["케이블 색상은 바뀔 수 있습니다.", "1번 핀이 3번 핀에 연결되는 케이블입니다."]
    source = (
        '<div class="se-component se-text"><p>앞 본문</p></div>'
        '<div class="se-component se-quotation se-l-quotation_line">'
        '<div class="se-module se-module-text"><blockquote class="se-quote">'
        f'<p><span>{quote_lines[0]}</span></p><p><span>{quote_lines[1]}</span></p>'
        '</blockquote></div></div>'
        '<div class="se-component se-text"><p>뒤 본문</p></div>'
    )
    with tempfile.TemporaryDirectory() as directory:
        with patch.dict(globals(), OUTPUT_ROOT=Path(directory)):
            preview_path, _, report = build_preview("test", {"originProduct": {"detailContent": source}})
        preview = preview_path.read_text(encoding="utf-8")
    assert '<blockquote class="quote-block">' in preview
    assert all(preview.count(line) == 1 for line in quote_lines)
    assert preview.index("앞 본문") < preview.index(quote_lines[0]) < preview.index(quote_lines[1]) < preview.index("뒤 본문")
    assert report["quotationComponentCount"] == 1
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
