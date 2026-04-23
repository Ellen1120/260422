"""
STM 문서 규칙 기반 파서 (API 불필요)
python-docx + 정규식으로 조제 정보·초자·시약량 추출
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from docx import Document

STM_FOLDER = Path(__file__).parent.parent / "STM"
DATA_FOLDER = Path(__file__).parent / "data"
KB_PATH = DATA_FOLDER / "knowledge_base.json"

# ── 정규식 상수 ───────────────────────────────────────────

_RE_KOREAN = re.compile(r"[가-힣]")

# 시험항목 키워드: 전체 줄이 일치해야 함 (줄 끝 앵커 $)
_RE_TEST_ITEM = re.compile(
    r"^(?:"
    r"Description"
    r"|Identification(?:\s+by\s+HPLC)?"
    r"|Dissolution"
    r"|Assay"
    r"|Uniformity\s+of\s+dosage\s+units?(?:\s*\([^)]*\))?"
    r"|Related\s+substances?(?:\s*\(Method\s+[AB]\))?"
    r"|Water\s+content\s+by\s+KF"
    r"|Polymorphism\s+by\s+PXRD"
    r"|Microbial\s+Enumeration\s+Test"
    r")\s*$",
    re.IGNORECASE,
)

# 조제 헤딩: "preparation" 단어를 포함하고 짧은 줄 (내용줄 제외)
_RE_PREP_CONTENT_START = re.compile(
    r"^(?:Accurately\s+)?(?:Weigh|Transfer|Add|Pipette|Dissolve"
    r"|Mix|Dilute|Note[:\s]|Centrifuge|Filter|Inject"
    r"|Proceed|Compare|Observe|Place|Take|Shake|Incubate"
    r"|Streak|Perform|After|The\s|If\s|Put\s|Pour\s|Blank"
    r"|Withdraw|Adjust|Heat|Cool|Rinse)",
    re.IGNORECASE,
)

# 함량 패턴 (standalone: "30 mg/25 mg", "50 mg")
_RE_STRENGTH_LINE = re.compile(
    r"^\s*(\d+(?:\.\d+)?\s*mg(?:\s*/\s*\d+(?:\.\d+)?\s*mg)*)\s*$",
    re.IGNORECASE,
)
# preparation (for X mg)
_RE_STRENGTH_PAREN = re.compile(
    r"\bpreparation\s*\(\s*for\s+(.+?)\s*\)",
    re.IGNORECASE,
)

# 시약 추출 패턴 (순서 중요: 구체적인 것 먼저)
_INGREDIENT_PATTERNS: list[tuple[str, re.Pattern]] = [
    # Weigh/Transfer X mg/g of [name] into/and/,
    ("weigh_transfer",
     re.compile(
         r"(?:Accurately\s+)?(?:Weigh\s+and\s+transfer|Transfer|Dissolve)"
         r"\s+(?:about\s+)?(\d+(?:\.\d+)?)\s*(mg|g)\s+of\s+"
         r"([A-Za-z0-9][^\n]{2,60}?)"
         r"(?:\s+(?:standard\b|into\b|to\b|in\b|and\b)|[,\.])",
         re.IGNORECASE,
     )),
    # Pipette X mL of [name]
    ("pipette",
     re.compile(
         r"Pipette\s+(\d+(?:\.\d+)?)\s*(mL|L)\s+of\s+"
         r"([A-Za-z0-9][^\n]{2,60}?)"
         r"(?:\s+(?:into\b|to\b|and\b)|[,\.])",
         re.IGNORECASE,
     )),
    # Add about X mL/mg/g of [name]
    ("add",
     re.compile(
         r"Add\s+(?:about\s+)?(\d+(?:\.\d+)?)\s*(mL|mg|g|L)\s+of\s+"
         r"([A-Za-z0-9][^\n]{2,60}?)"
         r"(?:\s+(?:and\b|to\b|,|\.))",
         re.IGNORECASE,
     )),
    # Dilute X mL of [substance] with water to Y mL  (희석)
    ("dilute",
     re.compile(
         r"Dilute\s+(\d+(?:\.\d+)?)\s*(mL|L)\s+of\s+"
         r"([A-Za-z0-9][^\n]{2,40}?)"
         r"\s+with\s+\w+\s+to\s+\d",
         re.IGNORECASE,
     )),
]

# Mix A and B in the ratio of X:Y v/v
_RE_RATIO = re.compile(
    r"Mix\s+(.+?)\s+and\s+(.+?)\s+in\s+the\s+ratio\s+of\s+"
    r"(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)",
    re.IGNORECASE,
)

# 초자 패턴
_RE_GLASSWARE = re.compile(
    r"(\d+(?:\.\d+)?)\s*mL\s+(?:of\s+)?(?:a\s+)?"
    r"(volumetric\s+flask|graduated\s+cylinder|measuring\s+cylinder"
    r"|pipette|burette|beaker|vial)",
    re.IGNORECASE,
)

# 용출 표준액 조제에서 volumetric flask 크기 추출
_RE_VOLUMETRIC_FLASK_SIZE = re.compile(
    r"(\d+(?:\.\d+)?)\s*mL\s+(?:of\s+)?volumetric\s+flask",
    re.IGNORECASE,
)

# 최종 볼륨 추출 (우선순위 순)
_RE_VOLUME_PATTERNS: list[re.Pattern] = [
    re.compile(r"dilute\s+to\s+(\d+(?:\.\d+)?)\s*mL", re.IGNORECASE),
    re.compile(r"with\s+\w[\w\s-]*?\s+to\s+(\d+(?:\.\d+)?)\s*mL", re.IGNORECASE),  # "with water to 1000 mL"
    re.compile(r"into\s+(?:a\s+)?(\d+(?:\.\d+)?)\s*mL\s+(?:of\s+)?volumetric\s+flask", re.IGNORECASE),
    re.compile(r"(\d+(?:\.\d+)?)\s*mL\s+volumetric\s+flask", re.IGNORECASE),
    re.compile(r"in\s+(\d+(?:\.\d+)?)\s*mL\s+of\s+(?:purified|Milli-Q|water)", re.IGNORECASE),
    re.compile(r"in\s+(\d+(?:\.\d+)?)\s*mL\s+of\s+\w", re.IGNORECASE),
]

# 제품명 추출
_RE_PRODUCT_NAME_WS = [
    re.compile(r"Ws\s*:.*?as\s+([A-Za-z][A-Za-z0-9\s\-]+?)(?:\s+sodium\b|\s+hydrate\b|\s+on\s+as|\s+\(|[,\.]|$)", re.IGNORECASE),
    re.compile(r"Ws\s*:\s*Weight\s+of\s+([A-Za-z][A-Za-z0-9\s\-]+?)\s+standard\s+taken", re.IGNORECASE),
    re.compile(r"Weigh\s+and\s+transfer\s+(?:about\s+)?[\d\.]+\s*(?:mg|g)\s+of\s+([A-Za-z][A-Za-z0-9\s\-]+?)\s+standard", re.IGNORECASE),
]


# ── 유틸 함수 ─────────────────────────────────────────────

def _is_korean(text: str) -> bool:
    return bool(_RE_KOREAN.search(text))


def _is_prep_heading(line: str) -> bool:
    """조제 섹션 헤딩 여부 판별."""
    if not re.search(r"\bpreparation\b", line, re.IGNORECASE):
        return False
    if len(line) > 120:
        return False
    if _RE_PREP_CONTENT_START.match(line):
        return False
    if line.strip().lower().startswith("note"):
        return False
    return True


def _derive_solution_name(heading: str) -> str:
    """'Buffer preparation' → 'Buffer'"""
    name = re.sub(r"\s+preparation\b", "", heading, flags=re.IGNORECASE).strip()
    return name.strip()


# ── 핵심 추출 함수 ────────────────────────────────────────

def _extract_volume(heading: str, lines: list[str]) -> float | None:
    """준비 텍스트에서 최종 조제 볼륨(mL) 추출."""
    combined = heading + " " + " ".join(lines)
    for pat in _RE_VOLUME_PATTERNS:
        m = pat.search(combined)
        if m:
            return float(m.group(1))
    return None


def _extract_ingredients(text: str, final_volume_ml: float | None) -> list[dict]:
    """준비 절차 텍스트에서 시약 목록 추출."""
    results: list[dict] = []
    seen: set[tuple] = set()

    def _add(name: str, amount: float, unit: str) -> None:
        name = name.strip().rstrip(",").strip()
        # 절차 설명 제거: ", sonicate for..." / ". sonicate..."
        name = re.split(r"[,\.]\s+(?:sonicate|mix|stir|shake|heat|cool|filter)", name, flags=re.IGNORECASE)[0].strip()
        # "standard" 접미사 제거
        name = re.sub(r"\s+standard$", "", name, flags=re.IGNORECASE).strip()
        # 앞에 붙은 "the " 제거
        name = re.sub(r"^the\s+", "", name, flags=re.IGNORECASE).strip()
        if len(name) < 3 or len(name) > 70:
            return
        # 모호한 참조 건너뜀
        if re.match(r"^(the\s+above|above|it|them|each|following)$", name, re.IGNORECASE):
            return
        key = (name.lower(), unit.lower())
        if key not in seen:
            seen.add(key)
            results.append({"name": name, "amount": round(amount, 4), "unit": unit.lower()})

    # 패턴 기반 추출
    for _tag, pat in _INGREDIENT_PATTERNS:
        for m in pat.finditer(text):
            _add(m.group(3), float(m.group(1)), m.group(2))

    # 비율 혼합 (Mix A and B in ratio X:Y)
    for m in _RE_RATIO.finditer(text):
        sub_a = m.group(1).strip()
        sub_b = m.group(2).strip()
        ratio_a = float(m.group(3))
        ratio_b = float(m.group(4))
        total = ratio_a + ratio_b
        if final_volume_ml and total > 0:
            _add(sub_a, round(ratio_a / total * final_volume_ml, 1), "mL")
            _add(sub_b, round(ratio_b / total * final_volume_ml, 1), "mL")

    return results


def _extract_glassware(text: str) -> list[dict]:
    """준비 절차 텍스트에서 초자 목록 추출."""
    results: list[dict] = []
    seen: set[tuple] = set()
    for m in _RE_GLASSWARE.finditer(text):
        size = m.group(1)
        gtype = re.sub(r"\s+", " ", m.group(2).lower().strip())
        key = (gtype, size)
        if key not in seen:
            seen.add(key)
            results.append({"type": gtype, "size": f"{size} mL", "count_per_batch": 1})
    return results


def _extract_product_names(paragraphs: list[str]) -> list[str]:
    """문서에서 제품명 추출 (복수 성분 지원)."""
    names: list[str] = []
    seen_lower: set[str] = set()

    for line in paragraphs:
        for pat in _RE_PRODUCT_NAME_WS:
            m = pat.search(line)
            if m:
                name = m.group(1).strip().rstrip(".,")
                # 숫자만 이거나 너무 짧으면 제외
                if re.match(r"^[\d\.\s]+$", name) or len(name) < 4:
                    continue
                # 불순물 표준품 제외
                if re.search(r"\bimpurit", name, re.IGNORECASE):
                    continue
                # 너무 긴 이름 제외 (변수 설명 등)
                if len(name) > 50:
                    continue
                # 제목화 (모두 소문자인 경우만)
                if name == name.lower():
                    name = name.title()
                if name.lower() not in seen_lower:
                    seen_lower.add(name.lower())
                    names.append(name)
    return names


def _extract_strengths(paragraphs: list[str]) -> list[str]:
    """문서에서 함량/규격 추출."""
    strengths: list[str] = []
    seen: set[str] = set()

    def _add(s: str) -> None:
        s = s.strip()
        if s and s not in seen:
            seen.add(s)
            strengths.append(s)

    for line in paragraphs:
        # "50 mg" or "30 mg/25 mg"
        m = _RE_STRENGTH_LINE.match(line)
        if m:
            _add(m.group(1).replace(" ", "").replace("mg", " mg").strip())
        # "preparation (for 50 mg)"
        m2 = _RE_STRENGTH_PAREN.search(line)
        if m2:
            _add(m2.group(1))

    # 단일 함량 문서: "Standard solution preparation (for 50 mg)" 패턴 없을 때
    # Label claim 수치로부터 추출 시도 (예: "20 mg Vonoprazan" 또는 section header)
    if not strengths:
        for line in paragraphs:
            m3 = re.search(
                r"(?:for|of)\s+(\d+(?:\.\d+)?\s*mg(?:\s*/\s*\d+(?:\.\d+)?\s*mg)*)\s+"
                r"(?:tablet|capsule|strength|dose)",
                line, re.IGNORECASE
            )
            if m3:
                _add(m3.group(1).strip())

    return strengths or ["N/A"]


# ── 섹션 파싱 ────────────────────────────────────────────

def _parse_sections(english_lines: list[str]) -> list[dict]:
    """
    영문 줄 목록을 시험항목 섹션으로 그룹화.
    반환: [{name, lines}]
    """
    sections: list[dict] = []
    current: dict | None = None

    for line in english_lines:
        if _RE_TEST_ITEM.match(line) and len(line.split()) <= 8:
            if current:
                sections.append(current)
            current = {"name": line.strip(), "lines": []}
        elif current is not None:
            current["lines"].append(line)

    if current:
        sections.append(current)
    return sections


def _parse_preparations(section_lines: list[str]) -> list[dict]:
    """
    섹션 내용 줄에서 조제 블록 추출.
    반환: [{section_name, solution_name, volume_per_batch_ml, preparation_text, ingredients, glassware}]
    """
    blocks: list[dict] = []
    current_heading: str | None = None
    current_lines: list[str] = []

    def _flush() -> None:
        nonlocal current_heading, current_lines
        if current_heading:
            prep_text = "\n".join(l for l in current_lines if l.strip())
            vol = _extract_volume(current_heading, current_lines)
            ingredients = _extract_ingredients(prep_text, vol)
            glassware = _extract_glassware(prep_text)
            blocks.append({
                "section_name": current_heading,
                "solution_name": _derive_solution_name(current_heading),
                "volume_per_batch_ml": vol,
                "preparation_text": prep_text,
                "ingredients": ingredients,
                "glassware": glassware,
            })
        current_heading = None
        current_lines = []

    for line in section_lines:
        if _is_prep_heading(line):
            _flush()
            current_heading = line.strip()
        elif current_heading is not None:
            current_lines.append(line)

    _flush()
    return blocks


# ── 용출 조건 추출 ───────────────────────────────────────

def _extract_dissolution_conditions(doc) -> dict | None:
    """
    문서 body 표에서 용출 조건(Volume, 배지, 장치, 속도) 파싱.
    Dissolution Media / Volume 두 행이 모두 있는 표를 탐색한다.
    """
    for tbl in doc.tables:
        row_map: dict[str, str] = {}
        for row in tbl.rows:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) >= 2:
                row_map[cells[0].strip()] = cells[1].strip()

        has_diss_media = any(
            re.match(r"dissolution\s+medi", k, re.IGNORECASE) for k in row_map
        )
        has_volume = any(re.match(r"volume\b", k, re.IGNORECASE) for k in row_map)
        if not (has_diss_media and has_volume):
            continue

        cond: dict = {"vessels_per_batch": 6}
        for raw_key, val in row_map.items():
            if re.match(r"dissolution\s+medi", raw_key, re.IGNORECASE):
                cond["medium_name"] = val
            elif re.match(r"volume\b", raw_key, re.IGNORECASE):
                m = re.search(r"(\d+(?:\.\d+)?)\s*mL", val, re.IGNORECASE)
                if m:
                    cond["volume_per_vessel_ml"] = float(m.group(1))
            elif re.match(r"usp\s+apparatus", raw_key, re.IGNORECASE):
                cond["apparatus"] = val
            elif re.match(r"speed\b", raw_key, re.IGNORECASE):
                ms = re.search(r"(\d+)", val)
                if ms:
                    cond["speed_rpm"] = int(ms.group(1))
            elif re.match(r"sampling\s+time", raw_key, re.IGNORECASE):
                cond["sampling_time"] = val

        if "volume_per_vessel_ml" in cond:
            return cond

    return None


def _extract_standard_medium_ml_for_dissolution(
    dissolution_section_lines: list[str],
    product_strengths: list[str],
) -> dict[str, float]:
    """
    용출 시험 섹션에서 함량별 표준액 조제 시 필요한 시험액(dissolution medium) 총량 추출.

    규칙:
    - 'standard'를 포함하지만 'sample'은 포함하지 않는 preparation 블록만 대상
    - dissolution medium을 희석제로 쓰는 블록 안의 volumetric flask 크기를 합산
    - 함량별 레이블(예: "30 mg/25 mg")이 블록 내에 있으면 함량별로 분리
    - 헤딩에 "(for X mg)" 패턴이 있으면 해당 함량에 귀속
    반환: {strength_label: total_ml}
    """
    # 준비 블록 분리
    blocks: list[tuple[str, list[str]]] = []
    cur_heading: str | None = None
    cur_lines: list[str] = []
    for line in dissolution_section_lines:
        if _is_prep_heading(line):
            if cur_heading is not None:
                blocks.append((cur_heading, cur_lines[:]))
            cur_heading = line.strip()
            cur_lines = []
        elif cur_heading is not None:
            cur_lines.append(line)
    if cur_heading is not None:
        blocks.append((cur_heading, cur_lines))

    per_strength: dict[str, float] = {}
    no_strength_total: float = 0.0

    for heading, lines in blocks:
        if not re.search(r"\bstandard\b", heading, re.IGNORECASE):
            continue
        if re.search(r"\bsample\b", heading, re.IGNORECASE):
            continue

        # 헤딩의 "(for X mg)" 패턴
        heading_strength: str | None = None
        m_hs = _RE_STRENGTH_PAREN.search(heading)
        if m_hs:
            heading_strength = m_hs.group(1).strip()

        # 블록 내 함량 레이블로 서브 블록 분리
        sub_blocks: list[tuple[str | None, list[str]]] = []
        cur_sub_str: str | None = None
        cur_sub_lines: list[str] = []
        found_str_in_lines = False

        for line in lines:
            m_sl = _RE_STRENGTH_LINE.match(line)
            if m_sl:
                if cur_sub_lines:
                    sub_blocks.append((cur_sub_str, cur_sub_lines[:]))
                cur_sub_str = m_sl.group(1).strip()
                cur_sub_lines = []
                found_str_in_lines = True
            else:
                cur_sub_lines.append(line)
        if cur_sub_lines:
            sub_blocks.append((cur_sub_str, cur_sub_lines))

        if not found_str_in_lines:
            sub_blocks = [(heading_strength, lines)]

        for sub_strength, sub_lines in sub_blocks:
            combined = heading + " " + " ".join(sub_lines)
            if not re.search(r"\bdissolution\s+(?:medium|media)\b", combined, re.IGNORECASE):
                continue

            # volumetric flask 크기 합산
            vol_sum = sum(
                float(m.group(1))
                for m in _RE_VOLUMETRIC_FLASK_SIZE.finditer(combined)
            )
            if vol_sum <= 0:
                vol_sum = _extract_volume(heading, sub_lines) or 0.0
            if vol_sum <= 0:
                continue

            if sub_strength:
                per_strength[sub_strength] = per_strength.get(sub_strength, 0.0) + vol_sum
            else:
                no_strength_total += vol_sum

    if per_strength:
        return per_strength
    if no_strength_total > 0:
        return {s: no_strength_total for s in product_strengths}
    return {}


# ── Reagent 표 파싱 ──────────────────────────────────────

def _extract_all_reagents(doc) -> list[dict]:
    """
    문서의 모든 Reagent 표(헤더 행: Reagent | Grade | Manufacturer | Cat. No. | Tracking No.)
    에서 시약 정보를 추출해 반환.
    """
    reagents: list[dict] = []
    seen: set[tuple] = set()

    for tbl in doc.tables:
        rows = [[c.text.strip() for c in row.cells] for row in tbl.rows]
        if not rows or len(rows[0]) < 2:
            continue

        # 첫 행이 Reagent 헤더인지 확인
        header = rows[0]
        if not re.match(r"^reagent\s*$", header[0], re.IGNORECASE):
            continue

        # 컬럼 인덱스 매핑
        col: dict[str, int] = {}
        for i, h in enumerate(header):
            hl = h.lower()
            if re.match(r"^reagent$", hl):
                col["name"] = i
            elif "tracking" in hl:
                col["tracking_no"] = i
            elif "grade" in hl:
                col["grade"] = i
            elif "manufacturer" in hl:
                col["manufacturer"] = i
            elif "cat" in hl:
                col["cat_no"] = i

        if "name" not in col or "tracking_no" not in col:
            continue

        for row in rows[1:]:
            name = row[col["name"]].strip() if col["name"] < len(row) else ""
            tracking = row[col["tracking_no"]].strip() if col["tracking_no"] < len(row) else ""
            if not name or not tracking:
                continue
            key = (name.lower(), tracking)
            if key in seen:
                continue
            seen.add(key)
            reagents.append({
                "name": name,
                "grade": row[col["grade"]].strip() if col.get("grade") is not None and col["grade"] < len(row) else "",
                "manufacturer": row[col["manufacturer"]].strip() if col.get("manufacturer") is not None and col["manufacturer"] < len(row) else "",
                "cat_no": row[col["cat_no"]].strip() if col.get("cat_no") is not None and col["cat_no"] < len(row) else "",
                "tracking_no": tracking,
            })

    return reagents


def _build_reagent_lookup(reagents: list[dict]) -> dict[str, list[str]]:
    """시약 이름(소문자) → tracking_no 리스트 (중복 제거)."""
    lookup: dict[str, list[str]] = {}
    for r in reagents:
        key = r["name"].lower()
        if key not in lookup:
            lookup[key] = []
        if r["tracking_no"] not in lookup[key]:
            lookup[key].append(r["tracking_no"])
    return lookup


def _enrich_ingredients_tracking(
    test_items: list[dict],
    reagent_lookup: dict[str, list[str]],
) -> None:
    """ingredients 각 항목에 tracking_no 필드를 추가 (일치하는 경우)."""
    for item in test_items:
        for prep in item.get("preparations", []):
            for ing in prep.get("ingredients", []):
                key = ing["name"].lower()
                tracking = reagent_lookup.get(key)
                if tracking:
                    ing["tracking_no"] = ", ".join(tracking)


# ── 헤더 표에서 제품명·함량 추출 ─────────────────────────

def _extract_header_title(doc) -> tuple[str | None, str | None]:
    """
    Word 섹션 헤더 표의 'Title: ...' 셀에서 (제품명, 함량)을 추출.
    예) 'Title: CTPH-D005 5 mg (Vonoprazan fumarate)'
         → ('CTPH-D005', '5 mg')
    예) 'Title: Pioglitazone Hydrochloride/Empagliflozin tablets'
         → ('Pioglitazone Hydrochloride/Empagliflozin tablets', None)
    """
    for section in doc.sections:
        for tbl in section.header.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    text = cell.text.strip()
                    m = re.match(r"Title:\s*(.+)", text, re.IGNORECASE)
                    if not m:
                        continue
                    # 한글 줄 제거: 첫 번째 줄(영문)만 사용
                    first_line = m.group(1).split("\n")[0].strip()
                    # 함량 추출: "CTPH-D005 5 mg (Vonoprazan fumarate)"
                    strength_m = re.search(
                        r"\b(\d+(?:\.\d+)?\s*mg(?:\s*/\s*\d+(?:\.\d+)?\s*mg)*)\b",
                        first_line, re.IGNORECASE
                    )
                    strength = strength_m.group(1).strip() if strength_m else None
                    # 제품명: 괄호 앞 부분, 함량 제거
                    name = re.sub(r"\s*\([^)]*\)\s*$", "", first_line).strip()
                    if strength:
                        # 함량 숫자를 제품명에서 제거 ("CTPH-D005 5 mg" → "CTPH-D005")
                        name = re.sub(
                            r"\s*\d+(?:\.\d+)?\s*mg(?:\s*/\s*\d+(?:\.\d+)?\s*mg)*\s*",
                            " ", name, flags=re.IGNORECASE
                        ).strip()
                    return name or None, strength
    return None, None


def _extract_doc_no(doc) -> str | None:
    """헤더 표에서 STM 문서 번호 추출 (예: STM-300599-T4)."""
    for section in doc.sections:
        for tbl in section.header.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    text = cell.text.strip()
                    m = re.match(r"^(STM-\S+)", text, re.IGNORECASE)
                    if m:
                        return m.group(1)
    return None


# ── 문서 파싱 진입점 ──────────────────────────────────────

def parse_document(doc_path: str) -> dict:
    """단일 STM .docx 파일을 파싱해 구조화된 dict 반환."""
    doc = Document(doc_path)

    # 헤더 표에서 제품명·함량 우선 추출
    header_name, header_strength = _extract_header_title(doc)

    # 영문 줄만 추출 (한글 줄 제외)
    english_lines = [
        p.text.strip()
        for p in doc.paragraphs
        if p.text.strip() and not _is_korean(p.text)
    ]

    # 제품명: 헤더 우선, 없으면 본문 파싱
    if header_name:
        product_name = header_name
    else:
        names = _extract_product_names(english_lines)
        product_name = " / ".join(names) if names else Path(doc_path).stem

    # 함량: 헤더에서 찾았으면 사용, 없으면 본문에서 추출
    if header_strength:
        strengths = [header_strength]
    else:
        strengths = _extract_strengths(english_lines)

    # 시험항목 섹션 파싱
    sections = _parse_sections(english_lines)
    test_items: list[dict] = []
    for sec in sections:
        preps = _parse_preparations(sec["lines"])
        test_items.append({"name": sec["name"], "preparations": preps})

    # 용출 시험 조건 + 표준액용 시험액 볼륨 추출
    diss_conditions = _extract_dissolution_conditions(doc)
    if diss_conditions:
        diss_sec_lines = next(
            (s["lines"] for s in sections
             if re.match(r"^dissolution\s*$", s["name"], re.IGNORECASE)),
            [],
        )
        diss_conditions["standard_medium_ml_by_strength"] = (
            _extract_standard_medium_ml_for_dissolution(diss_sec_lines, strengths)
        )
        for item in test_items:
            if re.match(r"^dissolution\s*$", item["name"], re.IGNORECASE):
                item["dissolution_conditions"] = diss_conditions
                break

    # Reagent 표 파싱 → ingredients에 tracking_no 추가
    reagents = _extract_all_reagents(doc)
    reagent_lookup = _build_reagent_lookup(reagents)
    _enrich_ingredients_tracking(test_items, reagent_lookup)

    # 문서 번호
    doc_no = _extract_doc_no(doc)

    return {
        "id": Path(doc_path).stem,
        "doc_no": doc_no,
        "stm_file": Path(doc_path).name,
        "product_name": product_name,
        "strengths": strengths,
        "test_items": test_items,
    }


def build_knowledge_base(progress_callback=None) -> list[dict]:
    """STM 폴더의 모든 .docx 파일을 파싱해 knowledge_base.json 저장."""
    products: list[dict] = []
    files = sorted(f for f in STM_FOLDER.glob("*.docx") if not f.name.startswith("~$"))

    for i, doc_path in enumerate(files):
        msg = f"Parsing {doc_path.name} ({i+1}/{len(files)})..."
        print(f"  → {msg}")
        if progress_callback:
            progress_callback(msg)
        product = parse_document(str(doc_path))
        products.append(product)

    DATA_FOLDER.mkdir(exist_ok=True)
    KB_PATH.write_text(
        json.dumps({"products": products}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Knowledge base saved → {KB_PATH}")
    return products


def load_knowledge_base() -> list[dict]:
    if KB_PATH.exists():
        data = json.loads(KB_PATH.read_text(encoding="utf-8"))
        return data.get("products", [])
    return []


def save_knowledge_base(products: list[dict]) -> None:
    DATA_FOLDER.mkdir(exist_ok=True)
    KB_PATH.write_text(
        json.dumps({"products": products}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
