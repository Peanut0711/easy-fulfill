"""상세 이미지 묶음을 외부 CSS 없이 같은 행과 비율로 렌더링한다."""

import html
import math
import re
from html.parser import HTMLParser


def normalized_widths(images, widths=None):
    def positive(value):
        try:
            number = float(value)
            return number if math.isfinite(number) and number > 0 else None
        except (TypeError, ValueError):
            return None

    values = [positive(value) for value in (widths or [])]
    if len(values) != len(images) or not all(values):
        values = []
        for image in images:
            width, height = positive(image.get("width")), positive(image.get("height"))
            values.append(width / height if width and height else 1)
    total = sum(values)
    return [value * 100 / total for value in values] if total else []


def render_image_group(images, widths=None):
    if not images:
        return ""
    weights = normalized_widths(images, widths)
    # 간격 셀도 백분율로 지정해 모바일에서 행 전체가 같은 비율로 축소된다.
    available = 100 - (len(images) - 1)
    cells = []
    for index, (image, weight) in enumerate(zip(images, weights)):
        if index:
            cells.append('<td aria-hidden="true" style="width:1%;padding:0;border:0;font-size:0"></td>')
        src = html.escape(image["src"], quote=True)
        alt = html.escape(image.get("alt", ""), quote=True)
        cells.append(
            f'<td style="width:{weight * available / 100:.6f}%;padding:0;border:0;vertical-align:top">'
            f'<img src="{src}" alt="{alt}" style="display:block;width:100%;max-width:100%;height:auto;margin:0;border:0">'
            '</td>'
        )
    return (
        f'<table data-ef-image-group="{len(images)}" role="presentation" '
        'style="width:100%;table-layout:fixed;border-collapse:collapse;border:0;margin:0 0 30px">'
        f'<tbody><tr>{"".join(cells)}</tr></tbody></table>'
    )


def upgrade_legacy_image_groups(source, image_sizes=None):
    """이전 grid 묶음도 이미지 크기 메타데이터가 있으면 비율을 복원한다."""
    class Images(HTMLParser):
        def __init__(self):
            super().__init__()
            self.images = []

        def handle_starttag(self, tag, attrs):
            if tag == "img":
                image = dict(attrs)
                if image.get("src"):
                    size = (image_sizes or {}).get(image["src"], {})
                    self.images.append({**image, **size})

    def replace(match):
        parser = Images()
        parser.feed(match[2])
        columns = int(match[1])
        return ''.join(render_image_group(parser.images[start:start + columns])
                       for start in range(0, len(parser.images), columns))

    return re.sub(r'<section class="image-block grid-([23])">(.*?)</section>', replace, source, flags=re.DOTALL)
