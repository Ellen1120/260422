"""
표준품 (Reference Standard) 조회 모듈
매칭 전략: 1) exact lowercase → 2) Excel 이름이 쿼리로 시작 → 3) stub
"""
from __future__ import annotations
import re
from pathlib import Path

_EXCEL_PATH = Path(__file__).parent.parent / "LIMS 도입 후 표준품 리스트_WS_최종본_250310.xlsx"

_DB: dict[str, list[dict]] = {}          # exact lowercase → entries
_ALL_ENTRIES: list[dict]    = []          # 전체 unique 항목 (prefix search용)


def _load() -> tuple[dict[str, list[dict]], list[dict]]:
    import openpyxl
    wb = openpyxl.load_workbook(_EXCEL_PATH, data_only=True, read_only=True)
    ws = wb.active
    header = [c.value for c in ws[1]]
    name_i = header.index("Name")
    ct_i   = header.index("Consumable Type")
    loc_i  = header.index("Location")

    db: dict[str, list[dict]] = {}
    seen_dedup: set[tuple] = set()
    all_entries: list[dict] = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        name = row[name_i]
        ct   = row[ct_i]
        loc  = row[loc_i]
        if not name:
            continue
        dedup_key = (str(name).strip(), str(ct or ""), str(loc or ""))
        if dedup_key in seen_dedup:
            continue
        seen_dedup.add(dedup_key)

        entry = {
            "name":            str(name).strip(),
            "consumable_type": str(ct or "").strip(),
            "location":        str(loc or "").strip(),
        }
        nk = entry["name"].lower()
        db.setdefault(nk, []).append(entry)
        all_entries.append(entry)

    wb.close()
    return db, all_entries


_ALIASES: dict[str, str] = {
    "scb4-impurity 1": "scb-4-imp-1",
}


def _find_matches(nm: str) -> list[dict]:
    """단일 이름에 대해 DB에서 매칭 엔트리 반환. 없으면 stub."""
    key = _ALIASES.get(nm.strip().lower(), nm.strip().lower())

    # 1. Exact match
    if key in _DB:
        return list(_DB[key])

    # 2. Prefix match: Excel 이름이 쿼리+공백 또는 쿼리 자체로 시작
    prefix_matches: list[dict] = []
    seen_dk: set[tuple] = set()
    for entry in _ALL_ENTRIES:
        ek = entry["name"].lower()
        if ek == key or ek.startswith(key + " "):
            dk = (entry["name"], entry["consumable_type"], entry["location"])
            if dk not in seen_dk:
                seen_dk.add(dk)
                prefix_matches.append(entry)
    if prefix_matches:
        return prefix_matches

    # 3. Stub — STM에 명시된 표준품이지만 Excel에 등록되지 않은 경우
    return [{"name": nm.strip(), "consumable_type": "", "location": ""}]


def lookup(names: list[str]) -> list[dict]:
    """STM에서 추출한 표준품 이름 목록을 조회해 반환."""
    global _DB, _ALL_ENTRIES
    if not _DB:
        try:
            _DB, _ALL_ENTRIES = _load()
        except Exception as e:
            print(f"[standards_db] Excel 로드 실패: {e}")
            return [{"name": nm.strip(), "consumable_type": "", "location": ""} for nm in names]

    out: list[dict] = []
    seen: set[tuple] = set()
    for nm in names:
        for m in _find_matches(nm):
            dk = (m["name"], m["consumable_type"], m["location"])
            if dk not in seen:
                seen.add(dk)
                out.append(m)
    return out
