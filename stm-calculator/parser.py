"""
STM 문서 규칙 기반 파서 (API 불필요)
python-docx + 정규식으로 조제 정보·초자·시약량 추출
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from docx import Document

# word namespace
_W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_W_T  = f'{{{_W_NS}}}t'
_W_TR = f'{{{_W_NS}}}tr'
_W_TC = f'{{{_W_NS}}}tc'

BASE_DIR = Path(__file__).resolve().parent
STM_DIR = BASE_DIR.parent / "STM"
STM_FOLDER = STM_DIR
DATA_FOLDER = BASE_DIR / "data"
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

# 한글 시험항목 이름 → 영문 정규 이름 매핑
_KO_TEST_ITEM_MAP: list[tuple[re.Pattern, str]] = [
    (re.compile(r'^성상\s*$'), 'Description'),
    (re.compile(r'^확인시험\s*(?:\(HPLC\))?\s*$'), 'Identification'),
    (re.compile(r'^용출\s*$'), 'Dissolution'),
    (re.compile(r'^함량\s*$'), 'Assay'),
    (re.compile(r'^제제\s*균일성\s*(?:\([^)]+\))?\s*$'), 'Uniformity of dosage units (Content Uniformity)'),
    (re.compile(r'^유연물질\s*(?:\(방법\s*[IVXivx가-힣]+\))?\s*$'), 'Related substances'),
    (re.compile(r'^수분\s*(?:함량|측정)?\s*(?:\(KF\))?\s*$'), 'Water content by KF'),
    (re.compile(r'^결정형\s*(?:\(PXRD\))?\s*$'), 'Polymorphism by PXRD'),
    (re.compile(r'^미생물\s*한도시험\s*$'), 'Microbial Enumeration Test'),
]


def _normalize_ko_strength(line: str) -> str | None:
    """'26 밀리그램/5 밀리그램' → '26 mg/5 mg', 한글 함량이 아니면 None."""
    if not _RE_KO_STRENGTH_LINE.match(line):
        return None
    nums = re.findall(r'\d+(?:\.\d+)?', line)
    return "/".join(f"{n} mg" for n in nums)


def _get_canonical_test_item_name(line: str) -> str | None:
    """영문 또는 한글 시험항목 이름이면 정규 영문명 반환, 아니면 None."""
    stripped = line.strip()
    if _RE_TEST_ITEM.match(stripped) and len(stripped.split()) <= 8:
        return stripped
    for pat, name in _KO_TEST_ITEM_MAP:
        if pat.match(stripped):
            return name
    return None


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
# 한글 함량 표기: "26 밀리그램/5 밀리그램"
_RE_KO_STRENGTH_LINE = re.compile(
    r"^\s*\d+(?:\.\d+)?\s*밀리그램(?:\s*/\s*\d+(?:\.\d+)?\s*밀리그램)*\s*$"
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

# 한글 시약 추출 패턴
_KO_INGREDIENT_PATTERNS: list[re.Pattern] = [
    # "아질사르탄 약 52 mg을 달아" / "Ammonium phosphate monobasic 약 2.0 g 달아"
    # "Potassium phosphate monobasic 6.81g과" (약 없음, 과 terminator)
    re.compile(
        r'([가-힣A-Za-z][가-힣A-Za-z0-9\s\-()]+?)\s+(?:약\s*)?(\d+(?:\.\d+)?)\s*(mg|g)\s*(?:[을를과]|달아)',
        re.IGNORECASE,
    ),
    # "아질사르탄 표준원액 5 mL 및" / "암로디핀베실산염 표준원액 3 mL을"
    re.compile(
        r'([가-힣A-Za-z][가-힣A-Za-z0-9\s\-()]+?)\s+(\d+(?:\.\d+)?)\s*(mL)\s*(?:[을를]|및|까지)',
        re.IGNORECASE,
    ),
]

# 한글 비율 혼합 패턴 (2성분: "A와 B를 X:Y v/v의 비율로 혼합")
_RE_RATIO_KO2 = re.compile(
    r'([가-힣A-Za-z][가-힣A-Za-z0-9\s\-()]*?)\s*(?:와|과)\s*'
    r'([가-힣A-Za-z][가-힣A-Za-z0-9\s\-()]*?)\s*(?:을|를)?\s*각각\s*'
    r'(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)',
)

# 초자 패턴 (영문)
_RE_GLASSWARE = re.compile(
    r"(\d+(?:\.\d+)?)\s*mL\s+(?:of\s+)?(?:a\s+)?"
    r"(volumetric\s+flask|graduated\s+cylinder|measuring\s+cylinder"
    r"|pipette|burette|beaker|vial)",
    re.IGNORECASE,
)
# 초자 패턴 (한글)
_RE_GLASSWARE_KO = re.compile(
    r"(\d+(?:\.\d+)?)\s*mL\s+(용량플라스크|메스실린더|메스플라스크|피펫|비이커|바이알)",
)

# 원심분리 팔콘 검출
_RE_CENTRIFUGE = re.compile(r"\bcentrifuge\b", re.IGNORECASE)

# 시린지 필터: 재질 명시형 (0.45 µm PVDF (Millipore))
_RE_FILTER_MENTION = re.compile(
    r"(\d+(?:\.\d+)?)\s*[uμµ]\s*m\s+(PVDF|Nylon|PTFE|PES|MCE|Cellulose)\s*(?:\(([^)]+)\))?",
    re.IGNORECASE,
)
# 시린지 필터: 재질 미명시형 (0.45µm syringe filter)
_RE_GENERIC_SYRINGE_FILTER = re.compile(
    r"(\d+(?:\.\d+)?)\s*[uμµ]\s*m\s+syringe\s+filter",
    re.IGNORECASE,
)
# 시린지 필터: 한글 표현 (0.45 µm PTFE 시린지 필터)
_RE_FILTER_KO = re.compile(
    r"(\d+(?:\.\d+)?)\s*[uμµ]\s*m\s+(PVDF|Nylon|PTFE|PES|MCE)?\s*시린지\s*필터",
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
    re.compile(r"with\s+\w[\w\s-]*?\s+to\s+(\d+(?:\.\d+)?)\s*mL", re.IGNORECASE),
    re.compile(r"into\s+(?:a\s+)?(\d+(?:\.\d+)?)\s*mL\s+(?:of\s+)?volumetric\s+flask", re.IGNORECASE),
    re.compile(r"(\d+(?:\.\d+)?)\s*mL\s+volumetric\s+flask", re.IGNORECASE),
    re.compile(r"into\s+(\d+(?:\.\d+)?)\s*mL\s+(?:of\s+)?(?:Milli-Q[-\s]?water|purified\s+water)", re.IGNORECASE),
    re.compile(r"in\s+(\d+(?:\.\d+)?)\s*mL\s+of\s+(?:purified|Milli-Q|water)", re.IGNORECASE),
    re.compile(r"in\s+(\d+(?:\.\d+)?)\s*mL\s+of\s+\w", re.IGNORECASE),
    # 한글: "X mL 용량플라스크에 넣고 표선" or "물 X mL에 넣고"
    re.compile(r"(\d+(?:\.\d+)?)\s*mL\s+용량플라스크"),
    re.compile(r"물\s+(\d+(?:\.\d+)?)\s*mL에\s+(?:넣|녹|옮)"),
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
    """조제 섹션 헤딩 여부 판별 (영문/한글 모두 지원)."""
    # 한글: "XX 조제" 패턴 (짧은 헤딩)
    if re.match(r'^[가-힣A-Za-z0-9\s()\-/]+\s*조제\s*$', line) and len(line) <= 40:
        return True
    # 영문: 기존 로직
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
    """'Buffer preparation' → 'Buffer', '완충액 조제' → '완충액'"""
    # 한글 "조제" 제거
    ko = re.sub(r'\s*조제\s*$', '', heading).strip()
    if ko != heading.strip():
        return ko
    # 영문 "preparation" 제거
    name = re.sub(r"\s+preparation\b", "", heading, flags=re.IGNORECASE).strip()
    return name.strip()


# ── 핵심 추출 함수 ────────────────────────────────────────

def _extract_volume(heading: str, lines: list[str]) -> float | None:
    """준비 텍스트에서 최종 조제 볼륨(mL) 추출. 명시적 볼륨 없으면 비율 합산으로 대체."""
    combined = heading + " " + " ".join(lines)
    for pat in _RE_VOLUME_PATTERNS:
        m = pat.search(combined)
        if m:
            return float(m.group(1))
    # 비율 혼합이고 총량 미지정인 경우 비율 합산을 기본 볼륨으로 사용
    m = _RE_RATIO.search(combined)
    if m:
        return float(m.group(3)) + float(m.group(4))
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

    # 패턴 기반 추출 (영문)
    for _tag, pat in _INGREDIENT_PATTERNS:
        for m in pat.finditer(text):
            _add(m.group(3), float(m.group(1)), m.group(2))

    # 한글 시약 추출
    for pat in _KO_INGREDIENT_PATTERNS:
        for m in pat.finditer(text):
            name_raw = m.group(1).strip().rstrip(',').strip()
            _add(name_raw, float(m.group(2)), m.group(3))

    # 비율 혼합 영문 (Mix A and B in ratio X:Y)
    for m in _RE_RATIO.finditer(text):
        sub_a = m.group(1).strip()
        sub_b = m.group(2).strip()
        ratio_a = float(m.group(3))
        ratio_b = float(m.group(4))
        total = ratio_a + ratio_b
        if final_volume_ml and total > 0:
            _add(sub_a, round(ratio_a / total * final_volume_ml, 1), "mL")
            _add(sub_b, round(ratio_b / total * final_volume_ml, 1), "mL")

    # 비율 혼합 한글 2성분 (A와 B를 각각 X:Y)
    for m in _RE_RATIO_KO2.finditer(text):
        sub_a = m.group(1).strip()
        sub_b = m.group(2).strip()
        ratio_a = float(m.group(3))
        ratio_b = float(m.group(4))
        total = ratio_a + ratio_b
        if final_volume_ml and total > 0:
            _add(sub_a, round(ratio_a / total * final_volume_ml, 1), "mL")
            _add(sub_b, round(ratio_b / total * final_volume_ml, 1), "mL")

    return results


_KO_GLASSWARE_MAP = {
    '용량플라스크': 'volumetric flask',
    '메스실린더': 'graduated cylinder',
    '메스플라스크': 'volumetric flask',
    '피펫': 'pipette',
    '비이커': 'beaker',
    '바이알': 'vial',
}

def _extract_glassware(text: str) -> list[dict]:
    """준비 절차 텍스트에서 초자 목록 추출. 동일 규격이 여러 번 나오면 count_per_batch 합산."""
    count_map: dict[tuple, int] = {}
    for m in _RE_GLASSWARE.finditer(text):
        size = m.group(1)
        gtype = re.sub(r"\s+", " ", m.group(2).lower().strip())
        key = (gtype, size)
        count_map[key] = count_map.get(key, 0) + 1
    for m in _RE_GLASSWARE_KO.finditer(text):
        size = m.group(1)
        gtype = _KO_GLASSWARE_MAP.get(m.group(2), m.group(2))
        key = (gtype, size)
        count_map[key] = count_map.get(key, 0) + 1
    return [
        {"type": gtype, "size": f"{size} mL", "count_per_batch": cnt}
        for (gtype, size), cnt in count_map.items()
    ]


def _extract_filters_from_text(text: str) -> list[dict]:
    """조제 텍스트에서 시린지 필터 추출 - 줄별로 처리해 복수 종류를 모두 포함."""
    filter_lines = [
        line for line in text.splitlines()
        if re.search(r"filter\s+through|syringe\s+filter|시린지\s*필터", line, re.IGNORECASE)
    ]
    if not filter_lines:
        return []

    results: list[dict] = []
    seen: set[tuple] = set()

    for line in filter_lines:
        specific_found = False
        for m in _RE_FILTER_MENTION.finditer(line):
            size = float(m.group(1))
            material = m.group(2).upper()
            mfr = m.group(3).strip() if m.group(3) else ""
            key = (size, material, mfr)
            if key not in seen:
                seen.add(key)
                results.append({
                    "size_um": size, "material": material,
                    "manufacturer": mfr, "filter_type": "syringe", "count_per_batch": 1,
                })
            specific_found = True
        if not specific_found:
            for m in _RE_GENERIC_SYRINGE_FILTER.finditer(line):
                size = float(m.group(1))
                key = (size, "", "")
                if key not in seen:
                    seen.add(key)
                    results.append({
                        "size_um": size, "material": "", "manufacturer": "",
                        "filter_type": "syringe", "count_per_batch": 1,
                    })
            for m in _RE_FILTER_KO.finditer(line):
                size = float(m.group(1))
                material = (m.group(2) or "").upper()
                key = (size, material, "")
                if key not in seen:
                    seen.add(key)
                    results.append({
                        "size_um": size, "material": material, "manufacturer": "",
                        "filter_type": "syringe", "count_per_batch": 1,
                    })
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

    # 한글 함량: "26 밀리그램/5 밀리그램"
    for line in paragraphs:
        ko_str = _normalize_ko_strength(line)
        if ko_str:
            _add(ko_str)

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

def _parent_section_name(name: str, current_name: str | None) -> str | None:
    """
    name이 current_name의 하위 항목이면 current_name 반환, 아니면 None.
    예) 'Identification by HPLC' → current 'Identification' → 'Identification'
    """
    if current_name and name.lower().startswith(current_name.lower() + " "):
        return current_name
    return None


def _parse_sections(english_lines: list[str]) -> list[dict]:
    """
    영문 줄 목록을 시험항목 섹션으로 그룹화.
    하위 항목(e.g. 'Identification by HPLC')은 상위 항목('Identification')에 병합.
    반환: [{name, lines}]
    """
    sections: list[dict] = []
    current: dict | None = None

    for line in english_lines:
        canonical = _get_canonical_test_item_name(line)
        if canonical:
            parent = _parent_section_name(canonical, current["name"] if current else None)
            # 같은 정규명 또는 하위 항목이면 현재 섹션에 병합
            if parent or (current and canonical == current["name"]):
                current["lines"].append(line)
            else:
                if current:
                    sections.append(current)
                current = {"name": canonical, "lines": []}
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
        if not current_heading:
            return

        sol_base = _derive_solution_name(current_heading)

        # current_lines 원소가 여러 줄을 포함할 수 있으므로 먼저 평탄화
        full_text = "\n".join(l for l in current_lines if l.strip())
        text_lines = full_text.splitlines()

        # 함량 서브섹션 분리: "30 mg/25 mg" 또는 "26 밀리그램/5 밀리그램" 줄이 있으면 각 함량별 prep 생성
        str_positions: list[tuple[int, str]] = []
        for i, l in enumerate(text_lines):
            m_en = _RE_STRENGTH_LINE.match(l)
            if m_en:
                str_positions.append((i, m_en.group(1).strip()))
                continue
            ko_str = _normalize_ko_strength(l)
            if ko_str:
                str_positions.append((i, ko_str))

        def _make_block(sol_name: str, sub_lines: list[str]) -> None:
            prep_text = "\n".join(l for l in sub_lines if l.strip())
            vol = _extract_volume(current_heading, sub_lines)
            ingredients = _extract_ingredients(prep_text, vol)
            glassware = _extract_glassware(prep_text)
            filters_list = _extract_filters_from_text(prep_text)
            if _RE_CENTRIFUGE.search(prep_text):
                filters_list.append({
                    "size_um": None, "material": None, "manufacturer": None,
                    "filter_type": "centrifuge", "count_per_batch": 1,
                })
            blocks.append({
                "section_name": current_heading,
                "solution_name": sol_name,
                "volume_per_batch_ml": vol,
                "preparation_text": prep_text,
                "ingredients": ingredients,
                "glassware": glassware,
                "filters": filters_list,
            })

        # 첫 번째 줄이 함량 헤더일 때만 분할 (Calculation 섹션 내 함량 참조 줄 오인식 방지)
        if str_positions and str_positions[0][0] == 0:
            for k, (start_idx, str_label) in enumerate(str_positions):
                end_idx = str_positions[k + 1][0] if k + 1 < len(str_positions) else len(text_lines)
                _make_block(f"{sol_base} (for {str_label})", text_lines[start_idx + 1:end_idx])
        else:
            _make_block(sol_base, text_lines)

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


# ── Body element 순서 반복 ────────────────────────────────

def _iter_body_elements(doc):
    """문서 body의 단락/표를 문서 순서대로 ('para', text) 또는 ('table', rows) yield."""
    for child in doc.element.body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            texts = [n.text or '' for n in child.iter() if n.tag == _W_T]
            yield 'para', ''.join(texts).strip()
        elif tag == 'tbl':
            rows: list[list[str]] = []
            for tr in child.iter(_W_TR):
                cells = []
                for tc in tr.iter(_W_TC):
                    ct = [n.text or '' for n in tc.iter(_W_T)]
                    cells.append(''.join(ct).strip())
                if cells:
                    rows.append(cells)
            if rows:
                yield 'table', rows


def _extract_hplc_conditions_per_section(doc) -> dict[str, dict]:
    """
    문서 body를 순서대로 스캔.
    각 시험항목 섹션에서 HPLC 조건 표(Flow rate + Run time)와
    직후 주입 순서 표(Solution | Injection No.)를 추출.
    반환: {section_name: {flow_rate_ml_min, run_time_min, injections:[...]}}
    """
    result: dict[str, dict] = {}
    current_section: str | None = None
    pending_hplc: dict | None = None

    for elem_type, elem_data in _iter_body_elements(doc):
        if elem_type == 'para':
            canonical = _get_canonical_test_item_name(elem_data)
            if canonical:
                current_section = canonical
                pending_hplc = None
            continue

        rows: list[list[str]] = elem_data
        if not current_section or not rows:
            continue

        row_map = {r[0]: r[1] for r in rows if len(r) >= 2}
        flow_key = next((k for k in row_map if re.match(r'^(?:flow\s*rate|유량)', k, re.IGNORECASE)), None)
        run_key  = next((k for k in row_map if re.match(r'^(?:run\s*time|분석\s*시간)', k, re.IGNORECASE)), None)

        col_key = next((k for k in row_map if re.match(r'^(?:column|칼럼)\s*$', k, re.IGNORECASE)), None)
        if flow_key and run_key and current_section not in result:
            fm = re.search(r'(\d+(?:\.\d+)?)', row_map[flow_key])
            rm = re.search(r'(\d+(?:\.\d+)?)', row_map[run_key])
            if fm and rm:
                pending_hplc = {
                    "flow_rate_ml_min": float(fm.group(1)),
                    "run_time_min":     float(rm.group(1)),
                    "column_spec": row_map[col_key].strip() if col_key else None,
                    "injections": [],
                }
            continue

        if pending_hplc is not None:
            header = rows[0]
            if (len(header) >= 2
                    and re.match(r'^solution', header[0], re.IGNORECASE)
                    and re.match(r'^injection', header[1], re.IGNORECASE)):
                for row in rows[1:]:
                    if len(row) < 2:
                        continue
                    # 각주 표기 제거: "2¹⁾" → "21)" → "2"
                    # 후미 단일 숫자+")" 패턴(각주 마커)만 제거
                    raw_count = re.sub(r'\d\)\s*$', '', row[1].strip()).strip()
                    count_m = re.search(r'(\d+)', raw_count) or re.search(r'(\d+)', row[1])
                    count   = int(count_m.group(1)) if count_m else 1
                    pending_hplc["injections"].append({
                        "solution":          row[0].strip(),
                        "count":             count,
                        "scales_with_batch": bool(re.search(r'\bsample\b|검액', row[0], re.IGNORECASE)),
                    })
                result[current_section] = pending_hplc
                pending_hplc = None

    return result


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
        # 한글 용출 조건 표: "용량" + "장치" 또는 "속도"
        has_ko_volume = any(re.match(r'^용량$', k) for k in row_map)
        has_ko_apparatus = any(re.match(r'^(?:장치|속도)', k) for k in row_map)
        is_ko_diss = has_ko_volume and has_ko_apparatus

        if not ((has_diss_media and has_volume) or is_ko_diss):
            continue

        cond: dict = {"vessels_per_batch": 6}
        for raw_key, val in row_map.items():
            # 영문 키
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
            # 한글 키
            elif re.match(r'^용량$', raw_key):
                m = re.search(r"(\d+(?:\.\d+)?)\s*mL", val, re.IGNORECASE)
                if m:
                    cond["volume_per_vessel_ml"] = float(m.group(1))
            elif re.match(r'^장치', raw_key):
                cond["apparatus"] = val
            elif re.match(r'^속도', raw_key):
                ms = re.search(r"(\d+)", val)
                if ms:
                    cond["speed_rpm"] = int(ms.group(1))
            elif re.match(r'^샘플링', raw_key):
                cond["sampling_time"] = val

        if is_ko_diss and not cond.get("medium_name"):
            cond["medium_name"] = "시험액"

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


# ── Standard 표 파싱 ─────────────────────────────────────

def _extract_standards_per_section(doc) -> dict[str, list[dict]]:
    """
    문서 body를 순서대로 스캔해 각 시험항목 섹션의 STD Name | Grade 표를 추출.
    반환: {section_name: [{"std_name": ..., "grade": ...}]}
    """
    result: dict[str, list[dict]] = {}
    current_section: str | None = None

    for elem_type, elem_data in _iter_body_elements(doc):
        if elem_type == 'para':
            canonical = _get_canonical_test_item_name(elem_data)
            if canonical:
                parent = _parent_section_name(canonical, current_section)
                if not parent:
                    current_section = canonical
            continue

        # 표 처리
        rows: list[list[str]] = elem_data
        if not current_section or not rows:
            continue

        header = rows[0]
        std_name_col = next(
            (i for i, h in enumerate(header) if re.match(r'(?:std\s*name|표준품)', h, re.IGNORECASE)),
            None,
        )
        if std_name_col is None:
            continue

        grade_col = next(
            (i for i, h in enumerate(header) if re.match(r'(?:grade|등급)', h, re.IGNORECASE)),
            None,
        )

        standards: list[dict] = []
        for row in rows[1:]:
            std_name = row[std_name_col].strip() if std_name_col < len(row) else ""
            grade = (
                row[grade_col].strip()
                if grade_col is not None and grade_col < len(row)
                else ""
            )
            if std_name:
                standards.append({"std_name": std_name, "grade": grade})

        if standards and current_section not in result:
            result[current_section] = standards

    return result


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

        # 첫 행이 Reagent 또는 한글 시약 헤더인지 확인
        header = rows[0]
        if not re.match(r"^(?:reagent|시약)\s*$", header[0], re.IGNORECASE):
            continue

        # 컬럼 인덱스 매핑 (영문/한글 공통)
        col: dict[str, int] = {}
        for i, h in enumerate(header):
            hl = h.lower()
            if re.match(r"^(?:reagent|시약)$", hl):
                col["name"] = i
            elif "tracking" in hl or "추적" in hl:
                col["tracking_no"] = i
            elif "grade" in hl or "등급" in hl:
                col["grade"] = i
            elif "manufacturer" in hl or "제조처" in hl:
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

    # 전체 단락 추출 (한글 문서 대응)
    all_lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    english_lines = [l for l in all_lines if not _is_korean(l)]

    # 제품명: 헤더 우선, 없으면 본문 파싱
    if header_name:
        product_name = header_name
    else:
        names = _extract_product_names(english_lines)
        product_name = " / ".join(names) if names else Path(doc_path).stem

    # 함량: 헤더에서 찾았으면 사용, 없으면 본문에서 추출 (한글 포함)
    if header_strength:
        strengths = [header_strength]
    else:
        strengths = _extract_strengths(all_lines)

    # 시험항목 섹션 파싱 (영문 + 한글 헤딩 모두 처리)
    sections = _parse_sections(all_lines)
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

    # 표준품(STD Name) 표 추출 → test_items에 연결
    standards_per_section = _extract_standards_per_section(doc)
    for item in test_items:
        if item["name"] in standards_per_section:
            item["standards"] = standards_per_section[item["name"]]

    # HPLC 크로마토그래피 조건 추출 (Mobile phase 볼륨 계산용)
    hplc_per_section = _extract_hplc_conditions_per_section(doc)
    for item in test_items:
        if item["name"] in hplc_per_section:
            item["hplc_conditions"] = hplc_per_section[item["name"]]

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
