"""여러 PC가 공유하는 일별 주문번호 인덱스(네이버·쿠팡·지마켓)의 병합 규칙.

각 PC는 마지막으로 확인한 시트 값(기준값)을 기억한다. 시트에 쓰거나 시트 값을 받을 때
스토어별로 다음 규칙을 적용해, 다른 PC의 변경을 덮어써서 번호가 되돌아가는 일을 막는다.

- 이 PC가 바꾸지 않은 스토어(로컬 = 기준값)는 시트 값을 따른다.
- 이 PC만 바꾼 스토어(시트 = 기준값)는 로컬 값을 쓴다.
- 양쪽이 모두 바꾼 스토어는 이미 두 PC가 번호를 썼을 수 있으므로 더 큰 값으로 맞추고 충돌로 알린다.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass

ORDER_INDEX_STORES = ("naver", "coupang", "gmarket")
STORE_NAMES = {"naver": "네이버", "coupang": "쿠팡", "gmarket": "지마켓"}


@dataclass(frozen=True)
class IndexConflict:
    store: str
    base: int | None
    local: int
    sheet: int
    merged: int

    def to_dict(self) -> dict:
        return asdict(self)


def merge_order_indices(local, base, sheet, *, local_dirty):
    """로컬·기준·시트 값을 병합해 (병합값 dict, 충돌 목록)을 반환한다.

    local_dirty가 False면(시트에 반영할 로컬 변경 없음) 시트 값을 그대로 따른다.
    base가 None이면 기준을 모르므로 로컬과 시트가 다른 스토어를 모두 충돌로 본다.
    """
    if not local_dirty:
        return {store: int(sheet[store]) for store in ORDER_INDEX_STORES}, []
    merged = {}
    conflicts = []
    for store in ORDER_INDEX_STORES:
        mine, theirs = int(local[store]), int(sheet[store])
        known = None if base is None else int(base[store])
        if mine == theirs:
            merged[store] = mine
        elif known is not None and mine == known:
            merged[store] = theirs
        elif known is not None and theirs == known:
            merged[store] = mine
        else:
            merged[store] = max(mine, theirs)
            conflicts.append(IndexConflict(store, known, mine, theirs, merged[store]))
    return merged, conflicts


def conflict_message(conflicts) -> str:
    """사용자 안내 문구: 겹쳤을 수 있는 번호 범위와 맞춘 값을 알려 준다."""
    lines = ["다른 PC가 같은 시간에 주문번호 인덱스를 바꿨습니다."]
    for item in conflicts:
        item = item if isinstance(item, IndexConflict) else IndexConflict(**item)
        name = STORE_NAMES.get(item.store, item.store)
        if item.base is not None:
            overlap_end = min(item.local, item.sheet) - 1
            overlap = (f"{item.base}~{overlap_end}번이 두 PC에서 함께 쓰였을 수 있습니다."
                       if overlap_end >= item.base else "겹친 번호는 없습니다.")
            lines.append(f"- {name}: 이 PC {item.base}→{item.local}, 다른 PC {item.base}→{item.sheet}. {overlap}")
        else:
            lines.append(f"- {name}: 이 PC {item.local}, 시트 {item.sheet} (마지막 동기화 값을 알 수 없음).")
        lines.append(f"  인덱스는 {item.merged}로 맞췄습니다.")
    lines.append("겹친 번호가 있으면 해당 주문의 번호를 확인해 주세요.")
    return "\n".join(lines)
