"""상세 본문의 원본 서식과 사이트 CSS에 의존하지 않는 기본 서식."""

import html
import re
from html.parser import HTMLParser


# 장식용 그림문자는 제거하되 사양의 ×, ≥, ±, °, Ω와 일반 글머리표는 유지한다.
_DECORATIVE_ICONS = re.compile(
    r"[\U0001f300-\U0001faff\U0001f1e6-\U0001f1ff\u2600-\u27bf]"
    r"(?:[\ufe0e\ufe0f\u200d\U000e0020-\U000e007f]*"
    r"[\U0001f300-\U0001faff\U0001f1e6-\U0001f1ff\u2600-\u27bf])*"
    r"[\ufe0e\ufe0f\u200d\U000e0020-\U000e007f]*[ \t]*"
)


def remove_decorative_icons(value):
    return _DECORATIVE_ICONS.sub("", value)


def source_text_style(attrs):
    """SmartEditor의 텍스트 관련 속성만 옮긴다. 임의 CSS/URL은 복사하지 않는다."""
    styles = {}
    for cls in attrs.get("class", "").split():
        size = re.fullmatch(r"se-fs-fs(\d+)", cls)
        align = re.fullmatch(r"se-text-paragraph-align-(left|center|right|justify)", cls)
        bullet = re.fullmatch(r"se-text-list-type-bullet-(disc|circle|square)", cls)
        if size and 6 <= int(size[1]) <= 96:
            styles["font-size"] = f"{size[1]}px"
        if align:
            styles["text-align"] = align[1]
        if bullet:
            styles["list-style-type"] = bullet[1]
    allowed = {
        "font-size": r"\d+(?:\.\d+)?(?:px|pt|em|rem|%)",
        "line-height": r"\d+(?:\.\d+)?(?:px|em|%)?|normal",
        "font-weight": r"[1-9]00|normal|bold",
        "font-style": r"normal|italic|oblique",
        "text-align": r"left|center|right|justify",
        "text-decoration": r"none|underline|line-through|underline line-through",
        "color": r"#[0-9a-f]{3,8}|[a-z]+|rgba?\([\d.,%\s]+\)",
        "background-color": r"#[0-9a-f]{3,8}|[a-z]+|rgba?\([\d.,%\s]+\)",
        "list-style-type": r"disc|circle|square|decimal|lower-alpha|upper-alpha|lower-roman|upper-roman|none",
    }
    for declaration in attrs.get("style", "").split(";"):
        key, sep, value = declaration.partition(":")
        key, value = key.strip().lower(), value.strip()
        if sep and key in allowed and re.fullmatch(allowed[key], value, re.I):
            styles[key] = value
    return ";".join(f"{key}:{value}" for key, value in styles.items())


def style_attribute(value):
    return f' style="{html.escape(value, quote=True)}"' if value else ""


class _PortableTextStyles(HTMLParser):
    """태그 위치만 수정해 이미지/영상/표 본문과 기존 속성은 그대로 보존한다."""

    def __init__(self, source, style_text=True):
        super().__init__(convert_charrefs=False)
        self.source = source
        self.style_text = style_text
        self.offsets = [0]
        self.offsets.extend(match.end() for match in re.finditer("\n", source))
        self.edits = []
        self.stack = []

    def handle_starttag(self, tag, attrs):
        attributes = dict(attrs)
        depth = sum(t in {"ul", "ol"} for t in self.stack)
        common = {
            "font-family": "inherit", "font-size": "inherit", "line-height": "inherit",
            "color": "inherit", "text-align": "inherit", "white-space": "normal",
            "overflow-wrap": "anywhere",
        }
        defaults = {}
        if tag in {"p", "ul", "ol", "li", "span", "strong", "b", "em", "i", "u", "s", "a"}:
            defaults.update(common)
            defaults["font-weight"] = "bold" if tag in {"strong", "b"} else "inherit"
            defaults["font-style"] = "italic" if tag in {"em", "i"} else "inherit"
        if tag == "p":
            defaults.update(display="block", margin="0" if "li" in self.stack else "0 0 12px", padding="0")
        elif tag in {"ul", "ol"}:
            marker = "decimal" if tag == "ol" else ("disc", "circle", "square")[min(depth, 2)]
            defaults.update(display="block", margin="4px 0" if depth else "8px 0 16px",
                            padding="0 0 0 1.5em", **{"list-style": f"{marker} outside"})
        elif tag == "li":
            defaults.update(display="list-item", margin="5px 0", padding="0", **{"list-style": "inherit"})
        elif tag in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            defaults.update(common)
            defaults.update(display="block", margin="0 0 20px", padding="0",
                            **{"font-size": "26px" if tag == "h1" else "24px", "line-height": "1.4",
                               "font-weight": "bold", "font-style": "normal", "text-align": "center"})
        if tag == "u":
            defaults["text-decoration"] = "underline"
        elif tag == "s":
            defaults["text-decoration"] = "line-through"
        elif tag == "a":
            defaults.update(color="#0655ba", **{"text-decoration": "underline"})
        if defaults and self.style_text:
            existing = attributes.get("style") or ""
            keys = {match[1].lower() for match in re.finditer(r"(?:^|;)\s*([\w-]+)\s*:", existing)}
            merged = ";".join(f"{key}:{value}" for key, value in defaults.items() if key not in keys)
            if existing:
                merged = merged + ";" + existing if merged else existing
            original = self.get_starttag_text()
            # 속성 전체를 파싱해 다른 속성값에 들어 있는 'style='을 잘못 바꾸지 않는다.
            rendered_attrs = "".join(
                f' {name}' + (f'="{html.escape(value, quote=True)}"' if value is not None else "")
                for name, value in attrs if name != "style"
            )
            replacement = f"<{tag}{rendered_attrs}{style_attribute(merged)}>"
            line, column = self.getpos()
            start = self.offsets[line - 1] + column
            self.edits.append((start, start + len(original), replacement))
        if tag not in {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr"}:
            self.stack.append(tag)

    def handle_endtag(self, tag):
        for index in range(len(self.stack) - 1, -1, -1):
            if self.stack[index] == tag:
                del self.stack[index:]
                break

    def handle_data(self, data):
        if any(tag in {"script", "style"} for tag in self.stack):
            return
        cleaned = remove_decorative_icons(data)
        if cleaned != data:
            line, column = self.getpos()
            start = self.offsets[line - 1] + column
            self.edits.append((start, start + len(data), cleaned))

    def handle_charref(self, name):
        # 기존 HTML의 숫자 참조로 작성된 아이콘도 본문에서 제거한다.
        if any(tag in {"script", "style"} for tag in self.stack):
            return
        token = f"&#{name};"
        decoded = html.unescape(token)
        if not remove_decorative_icons(decoded) or decoded in {"\ufe0e", "\ufe0f"}:
            line, column = self.getpos()
            start = self.offsets[line - 1] + column
            length = len(token) if self.source.startswith(token, start) else len(token) - 1
            self.edits.append((start, start + length, ""))


def portable_text_html(source, *, style_text=True):
    """본문 아이콘을 제거한다. 기존 편집본은 style_text=False로 서식을 보존한다."""
    parser = _PortableTextStyles(source, style_text=style_text)
    parser.feed(source)
    parser.close()
    for start, end, replacement in reversed(parser.edits):
        source = source[:start] + replacement + source[end:]
    return source
