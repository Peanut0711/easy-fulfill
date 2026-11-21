"""읽기 전용 상세 화면(QTextBrowser) 리치 텍스트 꾸밈 도우미: 카드·배지·구역 제목·라벨 표.

QTextBrowser 리치 텍스트는 둥근 모서리·한쪽 테두리를 지원하지 않아 표로 꾸민다.
카드 = 왼쪽 색 막대(좁은 칸) + 옅은 바탕 칸. 배지 = 바탕색을 칠한 굵은 글자. 색은 모두 테마 이름.
「문의」 탭(ui/main/inquiries.py)·「발주」 탭(ui/main/naver_orders.py)이 같이 쓴다. 값은 모두 이스케이프한다.
"""

import html

from shared import theme


def esc(value):
    return html.escape(str(value or ""))


def block(value):
    return esc(value).replace("\n", "<br>")


TONES = {  # 이름 → (글자·막대 색, 옅은 바탕)
    "primary": (theme.PRIMARY_PRESSED, theme.PRIMARY_SUBTLE),
    "info": (theme.INFO, theme.INFO_SUBTLE),
    "ok": (theme.OK, theme.OK_SUBTLE),
    "warn": (theme.WARN, theme.WARN_SUBTLE),
    "danger": (theme.DANGER, theme.DANGER_SUBTLE),
    "muted": (theme.TEXT_MUTED, theme.BG_PRESSED),
}


def chip(text, tone):
    fg, bg = TONES[tone]
    return (f'<span style="color:{fg}; background-color:{bg}; font-weight:600;">'
            f'&nbsp;{esc(text)}&nbsp;</span>')


def card(body, tone, plain_bg=False):
    """왼쪽 색 막대 카드. plain_bg면 바탕은 연한 회색(색은 막대만)."""
    accent, bg = TONES[tone]
    if plain_bg:
        bg = theme.BG_SUBTLE
    return (f'<table width="100%" cellspacing="0" cellpadding="0" style="margin-bottom:{theme.SPACE_SM}px;">'
            f'<tr><td width="4" bgcolor="{accent}"></td>'
            f'<td bgcolor="{bg}" style="padding:{theme.SPACE_SM}px {theme.SPACE_MD}px;">{body}</td></tr></table>')


def section(title, tone="muted", note=""):
    color = TONES[tone][0]
    extra = f' <span style="color:{theme.TEXT_MUTED}; font-weight:400;">{esc(note)}</span>' if note else ""
    return (f'<p style="margin-top:{theme.SPACE_LG}px; margin-bottom:{theme.SPACE_XS}px; '
            f'font-size:{theme.FONT_SIZE_SMALL}pt; font-weight:600; color:{color};">{esc(title)}{extra}</p>')


def facts(pairs):
    """흐린 라벨 · 값 두 열 표. pairs: [(라벨, 이미 이스케이프한 값 HTML)]."""
    rows = "".join(
        f'<tr><td style="color:{theme.TEXT_MUTED}; padding-right:{theme.SPACE_MD}px; white-space:nowrap;">'
        f'{esc(label)}</td><td>{value}</td></tr>'
        for label, value in pairs if value)
    if not rows:
        return ""
    return f'<table cellspacing="0" cellpadding="2">{rows}</table>'


def muted(text):
    return f'<span style="color:{theme.TEXT_MUTED};">{text}</span>'
