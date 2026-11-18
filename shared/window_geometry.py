"""메인 창 기본 크기·위치 계산과 저장된 위치 검사(Qt에 의존하지 않는 순수 함수).

기준은 docs/gui-window-settings-plan.md 「크기 기준」: FHD 100%에서 기본 900×800(프레임 제외),
작업 영역이 좁으면 폭 47%·높이 81% 안으로 줄인다. 사각형은 (x, y, 폭, 높이) 튜플이다.
"""

BASE_SIZE = (900, 800)
WIDTH_RATIO = 0.47
HEIGHT_RATIO = 0.81
# Windows 11 창 프레임(좌우 테두리 합, 제목 표시줄+아래 테두리) 추정값. 첫 표시 전에는 실제 값을
# 알 수 없어 Windows 11 100%에서 잰 값(frameGeometry − geometry = 0×31)을 쓴다.
FRAME_ESTIMATE = (0, 31)
# 저장된 위치를 되살릴 때 화면 안에 보여야 하는 최소 크기(창을 잡아 끌 수 있을 만큼)
VISIBLE_MIN = (200, 100)


def default_window_size(available, minimum, frame=FRAME_ESTIMATE, base=BASE_SIZE):
    """작업 영역 ``available``(폭, 높이)에 맞춘 기본 창 크기(프레임 제외)를 돌려준다.

    결과는 ``minimum``(폭, 높이)보다 작아지지 않는다.
    """
    avail_w, avail_h = available
    width = min(base[0], int(avail_w * WIDTH_RATIO) - frame[0])
    height = min(base[1], int(avail_h * HEIGHT_RATIO) - frame[1])
    return max(width, minimum[0]), max(height, minimum[1])


def centered_position(available_rect, size, frame=FRAME_ESTIMATE):
    """프레임을 포함한 창을 작업 영역 가운데에 둘 때의 창 내용 영역 왼쪽 위 좌표."""
    ax, ay, aw, ah = available_rect
    outer_w, outer_h = size[0] + frame[0], size[1] + frame[1]
    x = ax + max((aw - outer_w) // 2, 0) + frame[0] // 2
    y = ay + max((ah - outer_h) // 2, 0) + (frame[1] - frame[0] // 2)
    return x, y


def _intersection(a, b):
    left, top = max(a[0], b[0]), max(a[1], b[1])
    right, bottom = min(a[0] + a[2], b[0] + b[2]), min(a[1] + a[3], b[1] + b[3])
    return max(right - left, 0), max(bottom - top, 0)


def is_restorable(rect, screens, minimum):
    """저장된 창 영역 ``rect``를 그대로 되살려도 되는지 판단한다.

    - 크기가 ``minimum`` 이상이고, 어느 한 작업 영역(``screens``)에 들어가는 크기여야 한다.
    - 그 작업 영역 안에 ``VISIBLE_MIN`` 이상 보이고, 제목 표시줄이 작업 영역 위로 나가지 않아야 한다.
    모니터를 뺐거나 해상도·배율이 바뀌어 창이 화면 밖이나 화면보다 크게 뜨는 것을 막는다.
    """
    try:
        x, y, width, height = (int(v) for v in rect)
    except (TypeError, ValueError):
        return False
    if width < minimum[0] or height < minimum[1]:
        return False
    for screen in screens:
        sx, sy, sw, sh = screen
        if width + FRAME_ESTIMATE[0] > sw or height + FRAME_ESTIMATE[1] > sh:
            continue
        visible_w, visible_h = _intersection((x, y, width, height), screen)
        if visible_w >= VISIBLE_MIN[0] and visible_h >= VISIBLE_MIN[1] and y - FRAME_ESTIMATE[1] >= sy - 1:
            return True
    return False


def parse_saved_geometry(value):
    """앱 설정에 저장한 값({"x","y","w","h","maximized"})을 (사각형, 최대화) 로 바꾼다. 잘못되면 None."""
    if not isinstance(value, dict):
        return None
    try:
        rect = tuple(int(value[key]) for key in ("x", "y", "w", "h"))
    except (KeyError, TypeError, ValueError):
        return None
    return rect, bool(value.get("maximized", False))


def geometry_to_setting(rect, maximized):
    x, y, width, height = rect
    return {"x": int(x), "y": int(y), "w": int(width), "h": int(height), "maximized": bool(maximized)}


# 팝업 창(상세 HTML 편집기·이미지 보기): 원래 크기를 넘지 않고 작업 영역의 이 비율 안에 들어가게 한다.
DIALOG_AREA_RATIO = 0.9


def dialog_default_size(available, base, ratio=DIALOG_AREA_RATIO):
    """팝업 창 기본 크기(프레임 제외). ``base``(원래 크기)보다 크지 않고 작업 영역의 ``ratio`` 안에 든다."""
    avail_w, avail_h = available
    width = min(base[0], int(avail_w * ratio) - FRAME_ESTIMATE[0])
    height = min(base[1], int(avail_h * ratio) - FRAME_ESTIMATE[1])
    return max(width, 1), max(height, 1)
