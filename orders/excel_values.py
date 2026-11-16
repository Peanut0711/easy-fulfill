"""엑셀에서 읽은 셀 값을 비교·저장용 문자열로 바꾸는 공통 규칙."""

from __future__ import annotations

from shared.lazy_import import deferred_module

pd = deferred_module("pandas")


def excel_text(value) -> str:
    """빈 칸은 '', 정수 값을 가진 실수는 소수점 없이, 나머지는 앞뒤 공백을 뺀 문자열로 반환한다.

    pandas는 숫자 열에 빈 칸이 하나라도 있으면 열 전체를 실수형으로 읽어
    등기번호 6012345678901이 '6012345678901.0'이 된다. 번호·우편번호처럼 숫자로 저장된
    식별자를 문자열로 쓸 때 이 함수를 거친다.
    """
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()
