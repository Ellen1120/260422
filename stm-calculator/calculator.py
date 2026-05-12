"""
배치 수 기반 시험 자원 계산 엔진
"""
from __future__ import annotations

import math
import re

_RE_MP    = re.compile(r'^mobile\s*phase|^이동상', re.IGNORECASE)
_RE_BUF   = re.compile(r'^buffer\b|^완충액', re.IGNORECASE)
_RE_DILUENT = re.compile(r'^diluent|^희석액', re.IGNORECASE)
_RE_KOREAN = re.compile(r'[가-힣]')
_RE_SAMPLE_SOL  = re.compile(r'sample|검액', re.IGNORECASE)
_RE_UNIFORMITY  = re.compile(r'uniformity|균일성', re.IGNORECASE)
_RE_ASSAY       = re.compile(r'^assay\b', re.IGNORECASE)
# 표준액·표준원액·시스템 적합성: 배치 수·함량 수 무관 1회만 조제 (영문·한글 모두 포함)
_RE_STD_PREP    = re.compile(
    r'standard\s+(?:stock\s+)?solution|표준원액|표준액|시스템\s*적합성|system\s+suitability',
    re.IGNORECASE,
)
# 함량 지정 조제 패턴: "Standard solution (for 50 mg)"
_RE_FOR_STRENGTH = re.compile(r'\(\s*for\s+(.+?)\s*\)', re.IGNORECASE)
# 피펫 집계 대상: 검액 및 표준액만 포함 (용액 조제 시약류 제외 규칙 반영)
_RE_PIPETTE_SOURCE  = re.compile(r'sample|standard|표준원액|표준액|검액', re.IGNORECASE)

# Rule 3: 용액 조제 목록에서 제외할 시험 용액 (표준액, 표준원액, 검액, 시스템적합성 용액)
_RE_TEST_SOLUTION = re.compile(
    r'\bstandard\b.*\bsolution\b|\bsample\b|표준원액|표준액|검액'
    r'|system\s+suitability|시스템\s*적합성',
    re.IGNORECASE,
)
# Rule 1: 초자 목록에 포함할 소스 준비 (standard/sample/placebo/system suitability 준비에서 나온 초자만 유지)
_RE_GW_KEEP_SOURCE = re.compile(r'standard|sample|placebo|플라시보|표준|검액|system\s*suitability|시스템\s*적합성', re.IGNORECASE)

_RE_SAMPLE_OR_STD   = re.compile(r'sample|standard|검액|표준', re.IGNORECASE)
_RE_RATIO = re.compile(
    r'Mix\s+(.+?)\s+and\s+(.+?)\s+(?:in|at)\s+the\s+ratio\s+of\s+'
    r'(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)',
    re.IGNORECASE,
)
# 쉼표형 2성분 비율: "Mix A, B in/at the ratio of X:Y" (다단계 이동상 조제)
_RE_RATIO_COMMA = re.compile(
    r'Mix\s+(.+?)\s*,\s*(.+?)\s+(?:in|at)\s+the\s+ratio\s+of\s+'
    r'(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)'
    r'(?!\s*:\s*\d)',
    re.IGNORECASE,
)
# 영문 3성분 비율: "Mix A, B and C at/in the ratio of X:Y:Z"
_RE_RATIO_EN3 = re.compile(
    r'Mix\s+(.+?)\s*,\s*(.+?)\s+and\s+(.+?)\s+(?:in|at)\s+the\s+ratio\s+of\s+'
    r'(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)',
    re.IGNORECASE,
)
# Related substances Method A/B 패턴
_RE_RELATED_A = re.compile(r'Related\s+substances.*\(Method\s+A\)', re.IGNORECASE)
_RE_RELATED_B = re.compile(r'Related\s+substances.*\(Method\s+B\)', re.IGNORECASE)
# 한글 2성분 비율: "A와/과 B를/을 각각 X:Y"
_RE_RATIO_KO = re.compile(
    r'([가-힣A-Za-z][가-힣A-Za-z0-9\s\-()]*?)\s*(?:와|과)\s*'
    r'([가-힣A-Za-z][가-힣A-Za-z0-9\s\-()]*?)\s*(?:을|를)?\s*각각\s*'
    r'(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)',
)
# 한글 3성분 비율: "A, B 및 C를 X:Y:Z v/v/v"
_RE_RATIO_KO3 = re.compile(
    r'([가-힣A-Za-z][가-힣A-Za-z0-9\s\-()]*?)\s*,\s*'
    r'([가-힣A-Za-z][가-힣A-Za-z0-9\s\-()]*?)\s+및\s+'
    r'([가-힣A-Za-z][가-힣A-Za-z0-9\s\-()]*?)\s*(?:을|를)?\s*'
    r'(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)',
)
# 한글 2성분 비율 (및, 각각 없음): "A 및 B를 X:Y v/v"
_RE_RATIO_KO2_MIT = re.compile(
    r'([가-힣A-Za-z][가-힣A-Za-z0-9\s\-()]*?)\s+및\s+'
    r'([가-힣A-Za-z][가-힣A-Za-z0-9\s\-()]*?)\s*(?:을|를)?\s*'
    r'(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)'
    r'(?!\s*:\s*\d)',
)
# 한글 2성분 비율 (와/과 + 비율로 혼합): "A와 B를 X:Y (v/v)의 비율로 혼합"
_RE_RATIO_KO2_WA = re.compile(
    r'([가-힣A-Za-z][가-힣A-Za-z0-9\s\-()]*?)\s*(?:와|과)\s*'
    r'([가-힣A-Za-z][가-힣A-Za-z0-9\s\-()]*?)\s*(?:을|를)?\s*'
    r'(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)'
    r'(?!\s*:\s*\d)'
    r'(?:\s*\([vV]/[vV]\))?\s*(?:의)?\s*비율로',
)


# 국문 표준품명 → 영문 변환 사전
_KO_EN_STD: dict[str, str] = {
    "가바펜틴": "Gabapentin",
    "글리메피리드": "Glimepiride",
    "글리클라지드": "Gliclazide",
    "글리피지드": "Glipizide",
    "나프록센": "Naproxen",
    "네비볼롤": "Nebivolol",
    "니페디핀": "Nifedipine",
    "다파글리플로진": "Dapagliflozin",
    "덱시부프로펜": "Dexibuprofen",
    "독사조신메실산염": "Doxazosin mesylate",
    "둘록세틴염산염": "Duloxetine hydrochloride",
    "디클로페낙나트륨": "Diclofenac sodium",
    "라베프라졸나트륨": "Rabeprazole sodium",
    "라미프릴": "Ramipril",
    "레르카니디핀염산염": "Lercanidipine hydrochloride",
    "레보세티리진염산염": "Levocetirizine hydrochloride",
    "로사르탄칼륨": "Losartan potassium",
    "로수바스타틴칼슘": "Rosuvastatin calcium",
    "로수바스타틴": "Rosuvastatin",
    "리나글립틴": "Linagliptin",
    "리시노프릴": "Lisinopril",
    "메트포르민염산염": "Metformin hydrochloride",
    "메트포르민": "Metformin",
    "모사프리드시트르산염": "Mosapride citrate",
    "몬테루카스트나트륨": "Montelukast sodium",
    "미카르디스": "Telmisartan",
    "발사르탄": "Valsartan",
    "베나제프릴염산염": "Benazepril hydrochloride",
    "보노프라잔푸마르산염": "Vonoprazan fumarate",
    "보노프라잔": "Vonoprazan",
    "비소프롤롤푸마르산염": "Bisoprolol fumarate",
    "사쿠비트릴발사르탄나트륨수화물": "Sacubitril valsartan sodium hydrate",
    "사쿠비트릴": "Sacubitril",
    "서트랄린염산염": "Sertraline hydrochloride",
    "세티리진염산염": "Cetirizine hydrochloride",
    "소타롤염산염": "Sotalol hydrochloride",
    "시타글립틴인산염수화물": "Sitagliptin phosphate monohydrate",
    "시타글립틴": "Sitagliptin",
    "아모사핀": "Amoxapine",
    "아목시실린": "Amoxicillin",
    "아세클로페낙": "Aceclofenac",
    "아스피린": "Aspirin",
    "아질사르탄메독소밀": "Azilsartan medoxomil",
    "아질사르탄": "Azilsartan",
    "아토르바스타틴칼슘": "Atorvastatin calcium",
    "아토르바스타틴": "Atorvastatin",
    "아티카프릴": "Atazanavir",
    "알글리시다제알파": "Alglucosidase alfa",
    "암로디핀베실산염": "Amlodipine besylate",
    "암로디핀말레산염": "Amlodipine maleate",
    "암로디핀": "Amlodipine",
    "에소메프라졸마그네슘": "Esomeprazole magnesium",
    "에제티미브": "Ezetimibe",
    "에제티미브": "Ezetimibe",
    "엔탈라프릴말레산염": "Enalapril maleate",
    "엠파글리플로진": "Empagliflozin",
    "오메프라졸": "Omeprazole",
    "올메사르탄메독소밀": "Olmesartan medoxomil",
    "이르베사르탄": "Irbesartan",
    "이부프로펜": "Ibuprofen",
    "인다파미드": "Indapamide",
    "카나글리플로진": "Canagliflozin",
    "카르베딜롤": "Carvedilol",
    "카르베딜롤인산염": "Carvedilol phosphate",
    "칸데사르탄실렉세틸": "Candesartan cilexetil",
    "클로피도그렐황산염": "Clopidogrel bisulfate",
    "클로피도그렐": "Clopidogrel",
    "텔미사르탄": "Telmisartan",
    "트리메부틴말레산염": "Trimebutine maleate",
    "피오글리타존염산염": "Pioglitazone hydrochloride",
    "피오글리타존": "Pioglitazone",
    "하이드로클로로티아지드": "Hydrochlorothiazide",
    "히드로클로로티아지드": "Hydrochlorothiazide",
}


def _translate_std_name(name: str) -> str:
    """국문 표준품명을 영문으로 변환. 사전에 없으면 원문 반환."""
    return _KO_EN_STD.get(name.strip(), name)


def _sample_count_for_item(item_name: str) -> int:
    name = item_name.lower()
    if re.search(r'dissolution|용출', name):
        return 6
    if re.search(r'uniformity|균일성', name):
        return 10
    return 1


def _hplc_mp_volume(
    hplc: dict,
    batch_count: int,
    sample_count_per_batch: int | None = None,
) -> tuple[float, float, float, float | None, float | None, int | None, int] | None:
    """(rounded_ml, fixed_raw_ml, batch_raw_ml, nbf_raw, flow_rtime, total_samples, interval) 반환.
    sample_count_per_batch: 균일성 시험처럼 검액 수를 지정할 때 사용; 브라켓팅을 최대 20개당 1회(또는 지정된 간격)로 동적 계산."""
    flow  = hplc.get("flow_rate_ml_min")
    rtime = hplc.get("run_time_min")
    if not flow or not rtime:
        return None
    injs = hplc.get("injections", [])
    flow_rtime = flow * rtime
    interval = hplc.get("bracketing_interval", 20)

    if sample_count_per_batch is not None:
        # 검액 수 지정: 브라켓팅 동적 계산 (기본 최대 20개당 1회), 검액 주입 수 오버라이드
        total_samples = sample_count_per_batch * batch_count
        n_bracketings = math.ceil(total_samples / interval) + 1
        non_bracket_fixed_inj = sum(
            inj["count"] for inj in injs
            if not inj.get("scales_with_batch") and "bracketing" not in (inj.get("solution") or "").lower()
        ) if injs else 0
        fixed_inj = non_bracket_fixed_inj + n_bracketings
        batch_inj = sample_count_per_batch
        nbf_raw = flow * (rtime * non_bracket_fixed_inj + 30)
    else:
        fixed_inj = sum(inj["count"] for inj in injs if not inj.get("scales_with_batch")) if injs else 0
        batch_inj = sum(inj["count"] for inj in injs if inj.get("scales_with_batch")) if injs else 1
        total_samples = None
        nbf_raw = None

    fixed_raw = flow * (rtime * fixed_inj + 30)
    batch_raw = flow * rtime * batch_inj * batch_count
    raw = fixed_raw + batch_raw
    return math.ceil(raw / 50) * 50, fixed_raw, batch_raw, nbf_raw, flow_rtime, total_samples, interval


def _process_preparations(
    item_name: str,
    preparations: list[dict],
    batch_count: int,
    solutions: list[dict],
    glassware_agg: dict,
    filter_agg: dict,
    skip_sample: bool = False,
    strength: str | None = None,
):
    """단일 시험항목의 preparations를 순회하며 solutions/glassware/filter 집계."""
    sample_count = _sample_count_for_item(item_name)

    for prep in preparations:
        sol_name = prep.get("solution_name", prep.get("section_name", ""))

        # "(for 50 mg)" 처럼 함량 지정된 조제는 현재 함량만 포함
        # "for 80/10 mg, 40/10 mg" 처럼 복수 함량 쉼표/&/및 분리 대응
        if strength:
            m_str = _RE_FOR_STRENGTH.search(sol_name)
            if m_str:
                str_labels = [s.strip().lower() for s in re.split(r'\s*,\s*|\s+&\s+|\s+및\s+|\s+and\s+', m_str.group(1), flags=re.IGNORECASE)]
                if strength.strip().lower() not in str_labels:
                    continue

        # Blank 조제는 초자/필터 집계 대상에서도 완전히 제외
        if re.match(r'^blank\b', sol_name, re.IGNORECASE):
            continue

        is_sample_prep = bool(_RE_SAMPLE_SOL.search(sol_name))
        if skip_sample and is_sample_prep:
            continue

        # 표준액·표준원액, fixed_quantity 표시 조제는 배치 수에 무관하게 1회
        is_std_prep      = bool(_RE_STD_PREP.search(sol_name))
        effective_batches = 1 if (is_std_prep or prep.get("fixed_quantity")) else batch_count

        vol = prep.get("volume_per_batch_ml")
        # is_ratio_only: 비율 합산만으로 볼륨이 추정된 조제 → 이론량은 '-' 표시
        is_ratio_only = prep.get("is_ratio_only", False)
        theoretical = round(vol * effective_batches, 1) if (vol and not is_ratio_only) else None
        # ratio_ref_vol: CP025 NaOH 등 이론량 '-'이지만 성분 스케일링에 참조 볼륨 필요
        ratio_ref = prep.get("ratio_ref_vol")
        effective_vol = ratio_ref if (vol is None and ratio_ref) else vol

        # pH 조절용 용액: 이론량 '-', note 자동 추가 (용액명에 NaOH/HCl 포함 시)
        note_val = prep.get("note", "")
        _RE_PH_ADJ = re.compile(
            r'수산화나트륨|sodium\s+hydroxide|NaOH|염산|hydrochloric\s+acid|HCl|인산|phosphoric\s+acid',
            re.IGNORECASE,
        )
        if _RE_PH_ADJ.search(sol_name):
            if not note_val:
                note_val = "완충액 pH 조절용"
            theoretical = None  # pH 조절용은 이론량 표시 불필요

        exclude_from_solutions = prep.get("exclude_from_solutions", False)
        is_mobile_phase_flag = prep.get("is_mobile_phase", False)
        mp_fraction = prep.get("mobile_phase_fraction")  # gradient 비율 (0.0-1.0)

        sol_dict: dict = {
            "test_item": item_name,
            "solution_name": sol_name,
            "section_name": prep.get("section_name", ""),
            "volume_per_batch_ml": effective_vol,
            "theoretical_volume_ml": theoretical,
            "preparation_text": prep.get("preparation_text", ""),
            "ingredients": list(prep.get("ingredients", [])),
            "note": note_val,
        }
        if is_mobile_phase_flag:
            sol_dict["is_mobile_phase"] = True
        if mp_fraction is not None:
            sol_dict["_mp_fraction"] = mp_fraction
        if exclude_from_solutions:
            sol_dict["_excluded_from_output"] = True
        solutions.append(sol_dict)

        per_sample = sample_count if is_sample_prep else 1

        # 초자 집계 — 조제별/함량별/시험항목별 별도 행 (검액·표준액만 strength·test_item 포함)
        is_sample_std = bool(_RE_SAMPLE_OR_STD.search(sol_name))
        gw_strength  = (strength or "") if is_sample_std else ""
        gw_test_item = item_name        if is_sample_std else ""
        for gw in prep.get("glassware", []):
            key = (gw.get("type", ""), gw.get("size", ""), sol_name, gw_strength, gw_test_item)
            # mortar는 공유 장비이므로 배치 수·검액 수 무관하게 1개
            is_shared = gw.get("type", "") == "mortar"
            # Rule 18: 피펫은 배치 내 재사용 가능 → 검액 수 무관하게 1개
            is_pipette = gw.get("type", "") == "pipette"
            per_batch = 1 if (is_shared or is_pipette) else gw.get("count_per_batch", 1) * per_sample
            if key not in glassware_agg:
                glassware_agg[key] = {
                    "type": gw.get("type", ""),
                    "size": gw.get("size", ""),
                    "source_prep": sol_name,
                    "strength": gw_strength,
                    "test_item": gw_test_item,
                    "count_per_batch": per_batch,
                    # mortar: 공유 장비 → 총 1개. 피펫: 배치 내 재사용하나 배치마다 필요 → effective_batches개.
                    "total_count": 1 if is_shared else (effective_batches if is_pipette else per_batch * effective_batches),
                    # exclude_from_solutions 조제의 초자는 Rule 1 필터를 우회하여 목록에 포함
                    "from_excluded_solution": exclude_from_solutions,
                }
            else:
                if not is_shared:
                    glassware_agg[key]["count_per_batch"] += per_batch
                    glassware_agg[key]["total_count"] += per_batch * effective_batches

        # 필터 집계 — 조제별/시험항목별 별도 행 (검액·표준액만 strength·test_item 포함)
        fi_strength  = (strength or "") if is_sample_std else ""
        fi_test_item = item_name        if is_sample_std else ""
        for fi in prep.get("filters", []):
            mat   = fi.get("material") or ""
            mfr   = fi.get("manufacturer") or ""
            ftype = fi.get("filter_type", "syringe")
            size  = fi.get("size_um")   # None 허용 (centrifuge 팔콘)
            key = (size, mat, mfr, ftype, sol_name, fi_test_item)
            per_batch = 1 * per_sample
            if key not in filter_agg:
                filter_agg[key] = {
                    "size_um": size,
                    "material": mat,
                    "manufacturer": mfr,
                    "filter_type": ftype,
                    "source_prep": sol_name,
                    "strength": fi_strength,
                    "test_item": fi_test_item,
                    "count_per_batch": per_batch,
                    "total_count": per_batch * effective_batches,
                }
            else:
                filter_agg[key]["count_per_batch"] += per_batch
                filter_agg[key]["total_count"] += per_batch * effective_batches


def calculate_resources(
    product: dict,
    strength: str,
    test_item_names: list[str],
    batch_count: int,
) -> dict:
    solutions: list[dict] = []
    glassware_agg: dict[tuple, dict] = {}
    filter_agg: dict[tuple, dict] = {}

    # 균일성 시험항목별로 참조할 Assay 아이템을 assay_reference 필드 기반으로 매핑
    # {uniformity_item_name: assay_item}
    assay_map_for_uniformity: dict[str, dict] = {}
    has_uniformity = any(_RE_UNIFORMITY.search(n) for n in test_item_names)
    if has_uniformity:
        all_assay_items = [it for it in product.get("test_items", []) if _RE_ASSAY.match(it["name"])]
        for test_name in test_item_names:
            if not _RE_UNIFORMITY.search(test_name):
                continue
            uni_item = next((it for it in product.get("test_items", []) if it["name"] == test_name), None)
            ref = (uni_item or {}).get("assay_reference", "")
            matched = None
            if ref:
                # assay_reference로 정확 매핑
                matched = next((it for it in all_assay_items if ref in it["name"] or it["name"] in ref), None)
            if not matched and all_assay_items:
                matched = all_assay_items[0]
            if matched:
                assay_map_for_uniformity[test_name] = matched
    # 하위 호환: 단일 변수 (첫 번째 매핑값 사용)
    assay_item_for_uniformity = next(iter(assay_map_for_uniformity.values()), None)

    # Related substances (Method B)가 선택되고 Method A는 선택되지 않은 경우,
    # Method A의 공유 조제(완충액·이동상·희석액·표준액)를 상속
    related_a_for_b: dict | None = None
    has_related_b = any(_RE_RELATED_B.search(n) for n in test_item_names)
    has_related_a = any(_RE_RELATED_A.search(n) for n in test_item_names)
    if has_related_b and not has_related_a:
        for it in product.get("test_items", []):
            if _RE_RELATED_A.search(it["name"]):
                related_a_for_b = it
                break

    for item in product.get("test_items", []):
        if item["name"] not in test_item_names:
            continue

        _assay_ref = assay_map_for_uniformity.get(item["name"])
        if _RE_UNIFORMITY.search(item["name"]) and _assay_ref:
            # Uniformity: 해당 Assay의 비-sample 조제를 먼저 포함 (buffer, mobile phase, diluent 등)
            _process_preparations(
                item_name=item["name"],
                preparations=_assay_ref.get("preparations", []),
                batch_count=batch_count,
                solutions=solutions,
                glassware_agg=glassware_agg,
                filter_agg=filter_agg,
                skip_sample=True,
                strength=strength,
            )

        if _RE_RELATED_B.search(item["name"]) and related_a_for_b:
            # Method B: Method A의 비-sample, 비-placebo 공유 조제 먼저 포함
            shared_preps = [
                p for p in related_a_for_b.get("preparations", [])
                if not re.match(r'^placebo\b', p.get("solution_name", ""), re.IGNORECASE)
            ]
            _process_preparations(
                item_name=item["name"],
                preparations=shared_preps,
                batch_count=batch_count,
                solutions=solutions,
                glassware_agg=glassware_agg,
                filter_agg=filter_agg,
                skip_sample=True,
                strength=strength,
            )

        _process_preparations(
            item_name=item["name"],
            preparations=item.get("preparations", []),
            batch_count=batch_count,
            solutions=solutions,
            glassware_agg=glassware_agg,
            filter_agg=filter_agg,
            strength=strength,
        )

    # HPLC 기반 Mobile phase / Buffer 이론량 자동 계산
    # Uniformity는 Assay와 동일한 HPLC 조건을 사용하므로 Assay의 hplc_conditions 참조
    hplc_items = list(product.get("test_items", []))
    for item in hplc_items:
        item_name_for_hplc = item["name"]

        # Uniformity → 해당 Assay hplc_conditions 사용 (assay_reference 기반 per-item 매핑)
        if _RE_UNIFORMITY.search(item["name"]) and assay_map_for_uniformity.get(item["name"]):
            hplc = assay_map_for_uniformity[item["name"]].get("hplc_conditions")
            item_name_for_hplc = item["name"]
        # Method B → Method A hplc_conditions 사용 (Method A 미선택 시)
        elif _RE_RELATED_B.search(item["name"]) and related_a_for_b:
            hplc = related_a_for_b.get("hplc_conditions")
            item_name_for_hplc = item["name"]
        else:
            if item["name"] not in test_item_names:
                continue
            hplc = item.get("hplc_conditions")

        if not hplc:
            continue
        if item["name"] not in test_item_names:
            continue

        # 균일성 시험: 검액 10개/배치로 override, 브라켓팅 동적 계산
        _uni_sc = _sample_count_for_item(item["name"]) if _RE_UNIFORMITY.search(item["name"]) else None
        result = _hplc_mp_volume(hplc, batch_count, _uni_sc)
        if not result:
            continue
        mp_vol, hplc_fixed_raw, hplc_batch_raw, hplc_nbf_raw, hplc_flow_rtime, hplc_total_samp, hplc_interval = result

        # 이동상(이름 기준) + is_mobile_phase 플래그 & _mp_fraction 있는 용액(gradient 이동상)
        mp_sols = [
            s for s in solutions
            if s["test_item"] == item_name_for_hplc
            and (
                _RE_MP.match(s["solution_name"])
                or (s.get("is_mobile_phase") and s.get("_mp_fraction") is not None)
            )
        ]
        buf_sols = [s for s in solutions if s["test_item"] == item_name_for_hplc and _RE_BUF.match(s["solution_name"])]

        total_raw = hplc_fixed_raw + hplc_batch_raw  # gradient 비율 적용을 위한 반올림 전 총량

        total_buffer_needed = 0.0
        for sol in mp_sols:
            # gradient 비율이 있으면 해당 비율만큼만 할당, 없으면 전체 이동상 부피 사용
            frac = sol.get("_mp_fraction")
            sol_vol = math.ceil(total_raw * frac / 50) * 50 if frac is not None else mp_vol

            # 기존 조제 부피 (스케일링 전) 백업
            old_vol = sol.get("volume_per_batch_ml") or 1000.0
            sol["volume_per_batch_ml"]   = sol_vol
            sol["theoretical_volume_ml"] = sol_vol
            # merge에서 고정/배치 raw를 재합산할 수 있도록 저장
            sol["_hplc_fixed_raw"] = hplc_fixed_raw
            sol["_hplc_batch_raw"] = hplc_batch_raw
            if hplc_nbf_raw is not None:
                sol["_hplc_non_bracket_fixed_raw"] = hplc_nbf_raw
                sol["_hplc_flow_rtime"] = hplc_flow_rtime
                sol["_hplc_total_samples"] = hplc_total_samp
                sol["_hplc_interval"] = hplc_interval
            prep_text = sol.get("preparation_text", "")
            m3 = _RE_RATIO_EN3.search(prep_text) or _RE_RATIO_KO3.search(prep_text)
            if m3:
                sub_a, sub_b, sub_c = m3.group(1).strip(), m3.group(2).strip(), m3.group(3).strip()
                r_a, r_b, r_c = float(m3.group(4)), float(m3.group(5)), float(m3.group(6))
                t = r_a + r_b + r_c
                if t > 0:
                    new_ings = [
                        {"name": sub_a, "amount": round(r_a / t * sol_vol, 1), "unit": "mL"},
                        {"name": sub_b, "amount": round(r_b / t * sol_vol, 1), "unit": "mL"},
                        {"name": sub_c, "amount": round(r_c / t * sol_vol, 1), "unit": "mL"},
                    ]
                    # 기존 고체 성분(g, mg) 보존 및 스케일링
                    scale = sol_vol / old_vol
                    for old_ing in sol.get("ingredients", []):
                        if old_ing.get("unit", "").lower() in ["g", "mg"]:
                            scaled_ing = dict(old_ing)
                            scaled_ing["amount"] = round(float(old_ing["amount"]) * scale, 2)
                            new_ings.append(scaled_ing)
                    sol["ingredients"] = new_ings
                    for sub, r in [(sub_a, r_a), (sub_b, r_b), (sub_c, r_c)]:
                        if "buffer" in sub.lower() or "완충액" in sub:
                            total_buffer_needed += sol_vol * (r / t)
                            break
            else:
                m = _RE_RATIO.search(prep_text) or _RE_RATIO_KO.search(prep_text)
                if not m:
                    m = _RE_RATIO_KO2_MIT.search(prep_text) or _RE_RATIO_KO2_WA.search(prep_text)
                if not m:
                    m = _RE_RATIO_COMMA.search(prep_text)
                if m:
                    sub_a, sub_b = m.group(1).strip(), m.group(2).strip()
                    r_a, r_b = float(m.group(3)), float(m.group(4))
                    t = r_a + r_b
                    _above = re.compile(r'^(above|the\s+above|above\s+solution|above\s+mixture)$', re.IGNORECASE)
                    if t > 0:
                        new_ings = []
                        if not _above.match(sub_a):
                            new_ings.append({"name": sub_a, "amount": round(r_a / t * sol_vol, 1), "unit": "mL"})
                        if not _above.match(sub_b):
                            new_ings.append({"name": sub_b, "amount": round(r_b / t * sol_vol, 1), "unit": "mL"})
                        # 기존 고체 성분(g, mg) 보존 및 스케일링
                        scale = sol_vol / old_vol
                        for old_ing in sol.get("ingredients", []):
                            if old_ing.get("unit", "").lower() in ["g", "mg"]:
                                scaled_ing = dict(old_ing)
                                scaled_ing["amount"] = round(float(old_ing["amount"]) * scale, 2)
                                new_ings.append(scaled_ing)
                        sol["ingredients"] = new_ings

                        if "buffer" in sub_a.lower() or "완충액" in sub_a:
                            total_buffer_needed += sol_vol * (r_a / t)
                        elif "buffer" in sub_b.lower() or "완충액" in sub_b:
                            total_buffer_needed += sol_vol * (r_b / t)
                else:
                    # 비율 패턴 미매칭 시 mL/g/mg 성분을 sol_vol 기준으로 비례 스케일링
                    scale = sol_vol / old_vol
                    new_ings = []
                    for old_ing in sol.get("ingredients", []):
                        scaled_ing = dict(old_ing)
                        unit = scaled_ing.get("unit", "").lower()
                        if unit == "ml":
                            scaled_ing["amount"] = round(float(old_ing["amount"]) * scale, 1)
                            # ingredients에서 완충액 성분 비율로 total_buffer_needed 계산
                            ing_name = (old_ing.get("name") or "").lower()
                            if "buffer" in ing_name or "완충액" in ing_name:
                                total_buffer_needed += float(scaled_ing["amount"])
                        elif unit in ("g", "mg"):
                            scaled_ing["amount"] = round(float(old_ing["amount"]) * scale, 2)
                        new_ings.append(scaled_ing)
                    if new_ings:
                        sol["ingredients"] = new_ings

        if total_buffer_needed > 0 and buf_sols:
            total_mp_vol = sum(s.get("theoretical_volume_ml") or 0 for s in mp_sols) or mp_vol
            for sol in buf_sols:
                sol["theoretical_volume_ml"] = round(total_buffer_needed, 1)
                # merge에서 이동상 최종량 기준으로 완충액을 재계산할 수 있도록 비율 저장
                sol["_hplc_buf_ratio"] = total_buffer_needed / total_mp_vol if total_mp_vol else 0

        # gradient 이동상 성분으로 쓰인 중간 용액(is_mobile_phase 있으나 _mp_fraction 없음)의 이론량 파생
        # 예: 용액 A = 용액 B 성분(42%) + 용액 C 성분(20%)
        mp_ing_totals: dict[str, float] = {}
        for sol in mp_sols:
            for ing in sol.get("ingredients", []):
                if ing.get("unit", "").lower() == "ml":
                    k = (ing.get("name") or "").lower()
                    mp_ing_totals[k] = mp_ing_totals.get(k, 0) + float(ing.get("amount") or 0)
        if mp_ing_totals:
            for other_sol in solutions:
                if not other_sol.get("is_mobile_phase"):
                    continue
                if other_sol.get("_mp_fraction") is not None:
                    continue  # 직접 할당 이동상은 건너뜀
                if other_sol.get("theoretical_volume_ml"):
                    continue
                sname = (other_sol.get("solution_name") or "").lower()
                if sname in mp_ing_totals:
                    derived = round(mp_ing_totals[sname], 1)
                    other_sol["theoretical_volume_ml"] = derived
                    other_sol["volume_per_batch_ml"] = derived

    # 희석액/diluent: 다른 조제에서 사용된 총량 합산 → 이론량 + 비율 재료 업데이트
    _RE_FILL_DIL_IMPL  = re.compile(r'희석액으로\s*표선|diluent로\s*표선', re.IGNORECASE)
    _RE_FLASK_ML_IMPL  = re.compile(r'(\d+(?:\.\d+)?)\s*mL\s*(?:용량\s*플라스크|volumetric\s+flask)', re.IGNORECASE)
    _RE_ONCE_DIL_SOL   = re.compile(r'시스템\s*적합성|system\s*suitability', re.IGNORECASE)
    for sol in solutions:
        if not _RE_DILUENT.match(sol.get("solution_name", "")):
            continue
        if sol.get("theoretical_volume_ml"):
            continue
        sol_name_exact = sol["solution_name"]
        total_used = 0.0
        for other in solutions:
            if other is sol:
                continue
            other_vol   = other.get("volume_per_batch_ml") or 0
            other_theor = other.get("theoretical_volume_ml") or 0
            if other_vol:
                scale = other_theor / other_vol
            else:
                # vol 미정(None) 조제: 표준액은 1회, 그 외 검액 등은 batch_count배
                is_other_std = bool(_RE_STD_PREP.search(other.get("solution_name", "")))
                scale = 1.0 if is_other_std else float(batch_count)
            for ing in other.get("ingredients", []):
                if (ing.get("name") or "").lower() == sol_name_exact.lower() and \
                   ing.get("unit", "").lower() == "ml":
                    total_used += float(ing.get("amount", 0)) * scale
            # 희석액으로 표선까지 채우는 조제(표준액·시스템 적합성): explicit 성분 없으면 플라스크 볼륨으로 암묵적 사용량 추가
            other_text = other.get("preparation_text", "")
            if _RE_FILL_DIL_IMPL.search(other_text):
                has_explicit_dil = any(
                    (ing.get("name") or "").lower() == sol_name_exact.lower()
                    and ing.get("unit", "").lower() == "ml"
                    for ing in other.get("ingredients", [])
                )
                if not has_explicit_dil:
                    is_once = bool(
                        _RE_STD_PREP.search(other.get("solution_name", "")) or
                        _RE_ONCE_DIL_SOL.search(other.get("solution_name", ""))
                    )
                    impl_scale = 1.0 if is_once else float(batch_count)
                    for fm in _RE_FLASK_ML_IMPL.finditer(other_text):
                        total_used += float(fm.group(1)) * impl_scale
        if total_used > 0:
            prep_text = sol.get("preparation_text", "")
            sol["theoretical_volume_ml"] = round(total_used, 1)
            sol["volume_per_batch_ml"]   = total_used
            m3 = _RE_RATIO_EN3.search(prep_text) or _RE_RATIO_KO3.search(prep_text)
            m = None if m3 else (_RE_RATIO.search(prep_text) or _RE_RATIO_KO.search(prep_text) or _RE_RATIO_KO2_MIT.search(prep_text) or _RE_RATIO_KO2_WA.search(prep_text) or _RE_RATIO_COMMA.search(prep_text))
            if m3:
                sub_a, sub_b, sub_c = m3.group(1).strip(), m3.group(2).strip(), m3.group(3).strip()
                r_a, r_b, r_c = float(m3.group(4)), float(m3.group(5)), float(m3.group(6))
                t = r_a + r_b + r_c
                if t > 0:
                    sol["ingredients"] = [
                        {"name": sub_a, "amount": round(r_a / t * total_used, 1), "unit": "mL"},
                        {"name": sub_b, "amount": round(r_b / t * total_used, 1), "unit": "mL"},
                        {"name": sub_c, "amount": round(r_c / t * total_used, 1), "unit": "mL"},
                    ]
            elif m:
                sub_a = m.group(1).strip()
                sub_b = m.group(2).strip()
                r_a, r_b = float(m.group(3)), float(m.group(4))
                t = r_a + r_b
                if t > 0:
                    sol["ingredients"] = [
                        {"name": sub_a, "amount": round(r_a / t * total_used, 1), "unit": "mL"},
                        {"name": sub_b, "amount": round(r_b / t * total_used, 1), "unit": "mL"},
                    ]
            else:
                # 비율 패턴 미매칭: 기존 파싱된 성분 비율로 비례 스케일
                total_ing_ml = sum(
                    float(ing.get("amount", 0))
                    for ing in (sol.get("ingredients") or [])
                    if (ing.get("unit") or "").lower() == "ml"
                )
                if total_ing_ml > 0:
                    factor = total_used / total_ing_ml
                    sol["ingredients"] = [
                        {**ing, "amount": round(float(ing["amount"]) * factor, 1)}
                        if (ing.get("unit") or "").lower() == "ml" else ing
                        for ing in (sol.get("ingredients") or [])
                    ]
                else:
                    # 성분 미등록 조제: preparation_text 전체를 단일 용매명으로 사용
                    prep_text_clean = (prep_text or "").strip().rstrip(".")
                    if prep_text_clean:
                        sol["ingredients"] = [{"name": prep_text_clean, "amount": round(total_used, 1), "unit": "mL"}]

    # 희석액에 사용된 성분(용액 A 등)의 이론량을 희석액 이론량 기준으로 역산
    for dil_sol in solutions:
        if not _RE_DILUENT.match(dil_sol.get("solution_name", "")):
            continue
        dil_theor = dil_sol.get("theoretical_volume_ml") or 0
        dil_vol   = dil_sol.get("volume_per_batch_ml") or 0
        if not dil_theor or not dil_vol:
            continue
        for ing in dil_sol.get("ingredients", []):
            if (ing.get("unit") or "").lower() != "ml":
                continue
            ing_name   = (ing.get("name") or "").strip()
            ing_amount = float(ing.get("amount") or 0)
            if not ing_amount:
                continue
            ratio = ing_amount / dil_vol
            for other_sol in solutions:
                if other_sol is dil_sol:
                    continue
                if other_sol["solution_name"].strip().lower() == ing_name.lower():
                    other_sol["theoretical_volume_ml"] = round(ratio * dil_theor, 1)
                    break

    # 표준품 이름 수집 (STD Name 표 기반)
    rs_seen: set[str] = set()
    standard_names: list[str] = []
    for item in product.get("test_items", []):
        if item["name"] not in test_item_names:
            continue
        item_stds = item.get("standards", [])
        # Uniformity: 자체 standards 없으면 해당 Assay 표준품 사용
        if not item_stds and _RE_UNIFORMITY.search(item["name"]):
            _aref = assay_map_for_uniformity.get(item["name"]) or assay_item_for_uniformity
            if _aref:
                item_stds = _aref.get("standards", [])
        for std in item_stds:
            nm = _translate_std_name((std.get("std_name") or "").strip())
            if nm and nm.lower() not in rs_seen:
                rs_seen.add(nm.lower())
                standard_names.append(nm)

    dissolution_medium = _calc_dissolution_medium(
        product, strength, test_item_names, batch_count
    )

    # 완충액이 시험액(dissolution medium)인 경우 → 이론량을 총 dissolution medium으로 업데이트
    if dissolution_medium and dissolution_medium.get("medium_name"):
        medium_name_lower = dissolution_medium["medium_name"].lower().strip()
        for sol in solutions:
            sol_name_lower = sol.get("solution_name", "").lower()
            if medium_name_lower and (
                sol_name_lower.startswith(medium_name_lower) or
                medium_name_lower in sol_name_lower
            ):
                sol["theoretical_volume_ml"] = dissolution_medium["total_medium_ml"]
                sol["is_dissolution_medium"] = True

    # dissolution 표준액 prep 텍스트에서 완충액·희석액 추가 사용량 계산
    if dissolution_medium:
        _RE_FILL_DM_KO      = re.compile(r'시험액으로\s*표선|dissolution\s+medium으로\s*표선', re.IGNORECASE)
        _RE_FILL_DILUENT_KO = re.compile(r'희석액으로\s*표선|diluent로\s*표선', re.IGNORECASE)
        _RE_FLASK_ML        = re.compile(r'(\d+(?:\.\d+)?)\s*mL\s*(?:용량\s*플라스크|volumetric\s+flask)', re.IGNORECASE)
        for item in product.get("test_items", []):
            if item["name"] not in test_item_names:
                continue
            if not re.match(r"^dissolution\b", item["name"], re.IGNORECASE):
                continue
            for prep in item.get("preparations", []):
                pname = prep.get("solution_name", "")
                if not re.search(r'표준|standard', pname, re.IGNORECASE):
                    continue
                prep_text = prep.get("preparation_text", "")
                flask_sizes = [float(m.group(1)) for m in _RE_FLASK_ML.finditer(prep_text)]
                if not flask_sizes:
                    continue
                # 시험액(dissolution medium)으로 표선 → 표준액 dissolution medium 사용량 추가
                if _RE_FILL_DM_KO.search(prep_text) and dissolution_medium.get("standard_medium_ml_once", 0) == 0:
                    std_vol = max(flask_sizes)
                    dissolution_medium["standard_medium_ml_once"] = std_vol
                    dissolution_medium["total_medium_ml"] = round(
                        dissolution_medium["sample_medium_ml"] + std_vol, 1
                    )
                    for sol in solutions:
                        if sol.get("is_dissolution_medium"):
                            sol["theoretical_volume_ml"] = dissolution_medium["total_medium_ml"]
                # 희석액으로 표선 → 희석액 이론량 설정 (표준액 조제는 1회이므로 batch_count 미적용)
                if _RE_FILL_DILUENT_KO.search(prep_text):
                    dil_flask_vol = min(flask_sizes)
                    for sol in solutions:
                        sn = sol.get("solution_name", "").lower()
                        if ("희석액" in sn or "diluent" in sn) and sol.get("volume_per_batch_ml"):
                            if sol.get("theoretical_volume_ml") is None:
                                sol["theoretical_volume_ml"] = round(dil_flask_vol, 1)
                            break

    # pipette: mL 재료 규격 추출 → 초자 목록(glassware_agg)에 통합
    _RE_BTV = re.compile(
        r'(?:dilute to volume|bring to volume|make up to volume|make to volume)\s+with\s+([^.,\n]+)',
        re.IGNORECASE,
    )
    for sol in solutions:
        if not _RE_PIPETTE_SOURCE.search(sol.get("solution_name", "")):
            continue

        prep_text = sol.get("preparation_text", "")
        # 비율 혼합 용액(Mix A:B)은 pipette 자동 생성 제외
        if (_RE_RATIO.search(prep_text) or _RE_RATIO_KO.search(prep_text) or
                _RE_RATIO_KO3.search(prep_text) or _RE_RATIO_KO2_MIT.search(prep_text) or
                _RE_RATIO_KO2_WA.search(prep_text) or _RE_RATIO_COMMA.search(prep_text)):
            continue
        btv_m = _RE_BTV.search(prep_text)
        btv_solvent = btv_m.group(1).strip().lower() if btv_m else ""
        vol_per_batch = sol.get("volume_per_batch_ml") or 0
        for ing in sol.get("ingredients", []):
            if ing.get("unit", "").strip().lower() != "ml":
                continue
            f_vol = float(ing.get("amount", 0))
            if f_vol <= 0 or f_vol > 25:
                continue
            
            # 규격 정규화 (10.0 -> 10)
            if f_vol == int(f_vol):
                size = f"{int(f_vol)} mL"
            else:
                size = f"{f_vol} mL"
                
            ing_name = ing.get("name", "").strip().lower()
            # 표선 용매와 동일한 성분이면 피펫 불필요
            if btv_solvent and (ing_name in btv_solvent or btv_solvent.startswith(ing_name)):
                continue
            sol_name = sol.get("solution_name", "")
            sol_item = sol.get("test_item", "")
            is_ss = bool(_RE_SAMPLE_OR_STD.search(sol_name))
            pip_strength  = (strength or "") if is_ss else ""
            pip_test_item = sol_item if is_ss else ""
            # 표준액은 배치 무관 1회, 검액은 배치마다 필요
            pip_batches = 1 if bool(_RE_STD_PREP.search(sol_name)) else batch_count
            key  = ("pipette", size, sol_name, pip_strength, pip_test_item)
            if key not in glassware_agg:
                glassware_agg[key] = {
                    "type": "pipette",
                    "size": size,
                    "source_prep": sol_name,
                    "strength": pip_strength,
                    "test_item": pip_test_item,
                    "count_per_batch": 1,
                    "total_count": pip_batches,
                }

    # "X mL를 취해/취하여" 패턴 → 피펫 (위 액 취하기: 표준액·검액 제조 시)
    _RE_TAKE_KO = re.compile(r'(\d+(?:\.\d+)?)\s*mL\s*를?\s*(?:취해|취하여)', re.IGNORECASE)
    for sol in solutions:
        if not _RE_PIPETTE_SOURCE.search(sol.get("solution_name", "")):
            continue
        prep_text = sol.get("preparation_text", "")
        for m in _RE_TAKE_KO.finditer(prep_text):
            f_vol = float(m.group(1))
            if f_vol <= 0 or f_vol > 25:
                continue
            size = f"{int(f_vol)} mL" if f_vol == int(f_vol) else f"{f_vol} mL"
            sol_name = sol.get("solution_name", "")
            sol_item = sol.get("test_item", "")
            is_ss = bool(_RE_SAMPLE_OR_STD.search(sol_name))
            pip_batches = 1 if bool(_RE_STD_PREP.search(sol_name)) else batch_count
            key = ("pipette", size, sol_name, (strength or "") if is_ss else "", sol_item if is_ss else "")
            if key not in glassware_agg:
                glassware_agg[key] = {
                    "type": "pipette",
                    "size": size,
                    "source_prep": sol_name,
                    "strength": (strength or "") if is_ss else "",
                    "test_item": sol_item if is_ss else "",
                    "count_per_batch": 1,
                    "total_count": pip_batches,
                }

    # exclude_from_solutions 조제: 희석액 계산 참여 후 솔루션 목록에서 제거
    solutions = [s for s in solutions if not s.get("_excluded_from_output")]

    # 내용 없는 빈 조제(이중언어 문서에서 영문 헤딩 stub) 제거
    solutions = [
        s for s in solutions
        if s.get("preparation_text") or s.get("ingredients")
    ]

    # Rule 3: 표준액, 표준원액, 검액은 용액 조제 목록에서 제외
    solutions = [
        s for s in solutions
        if not _RE_TEST_SOLUTION.search(s.get("solution_name", ""))
    ]

    # Blank 용액은 조제 목록에서 제외
    solutions = [
        s for s in solutions
        if not re.match(r'^blank\b', s.get("solution_name", ""), re.IGNORECASE)
    ]

    # 동일 용액 내 중복 성분 제거: 추적번호 있는 영문명 우선, 동일 양/단위의 한글명 제거
    for sol in solutions:
        ings = sol.get("ingredients")
        if not ings or len(ings) < 2:
            continue
        seen: dict[tuple, dict] = {}
        cleaned: list[dict] = []
        for ing in ings:
            key = (round(float(ing.get("amount", 0)), 4), (ing.get("unit") or "").lower())
            if key in seen:
                existing = seen[key]
                if existing.get("tracking_no") and _RE_KOREAN.search(ing.get("name", "")):
                    continue  # 추적번호 있는 기존 유지, 한글명 중복 제거
                if ing.get("tracking_no") and _RE_KOREAN.search(existing.get("name", "")):
                    cleaned.remove(existing)
                    seen[key] = ing
                    cleaned.append(ing)
                    continue
            else:
                seen[key] = ing
            cleaned.append(ing)
        sol["ingredients"] = cleaned

    # Rule 1: 용액 조제(시약)에서 나온 초자는 초자 목록에서 제외 — standard/sample 준비 초자만 유지
    # exclude_from_solutions 표시 조제(예: Atorvastatin compound 용액)의 초자는 예외적으로 포함
    glassware_list = [
        g for g in glassware_agg.values()
        if _RE_GW_KEEP_SOURCE.search(g.get("source_prep", ""))
        or g.get("from_excluded_solution")
    ]

    # Rule 7: 용출 시험 시 900mL 피펫은 불필요 (용출 배지를 피펫으로 옮기지 않음)
    glassware_list = [
        g for g in glassware_list
        if not (g.get("type") == "pipette" and "900" in str(g.get("size", "")))
    ]

    # Rule 14: 이동상(또는 is_mobile_phase 플래그) → 완충액 순; 희석액은 맨 뒤
    _is_mp = lambda s: _RE_MP.match(s.get("solution_name", "")) or s.get("is_mobile_phase")
    _mp  = [s for s in solutions if _is_mp(s)]
    _buf = [s for s in solutions if _RE_BUF.match(s.get("solution_name", "")) and not _is_mp(s)]
    _dil = [s for s in solutions if _RE_DILUENT.match(s.get("solution_name", ""))
            and not _is_mp(s)
            and not _RE_BUF.match(s.get("solution_name", ""))]
    _other = [s for s in solutions
              if not _is_mp(s)
              and not _RE_BUF.match(s.get("solution_name", ""))
              and not _RE_DILUENT.match(s.get("solution_name", ""))]
    solutions = _mp + _buf + _dil + _other

    return {
        "product_name": product.get("product_name", ""),
        "doc_no": product.get("doc_no"),
        "stm_file": product.get("stm_file", ""),
        "strength": strength,
        "test_items": test_item_names,
        "batch_count": batch_count,
        "solutions": solutions,
        "glassware": glassware_list,
        "filters": list(filter_agg.values()),
        "pipettes": [],
        "dissolution_medium": dissolution_medium,
        "standard_names": standard_names,
    }


def _calc_dissolution_medium(
    product: dict,
    strength: str,
    test_item_names: list[str],
    batch_count: int,
) -> dict | None:
    for item in product.get("test_items", []):
        if item["name"] not in test_item_names:
            continue
        if not re.match(r"^dissolution\b", item["name"], re.IGNORECASE):
            continue

        conds = item.get("dissolution_conditions")
        if not conds:
            return None

        vol_per_vessel = conds.get("volume_per_vessel_ml", 0)
        if not vol_per_vessel:
            return None

        vessels  = conds.get("vessels_per_batch", 6)
        std_by_str = conds.get("standard_medium_ml_by_strength", {})
        # Rule 15: 정확 매칭 후 실패 시 쉼표/& 구분 복합 키에서 부분 매칭
        # Rule 17: 같은 표준액 조제가 복수 함량에 공유될 때 중복 집계 방지를 위해 키 반환
        std_once = 0.0
        matched_std_key = ""
        if std_by_str:
            if strength in std_by_str:
                std_once = float(std_by_str[strength])
                matched_std_key = strength
            else:
                for key, val in std_by_str.items():
                    if strength in [s.strip() for s in re.split(r',|&|및', key)]:
                        std_once = float(val)
                        matched_std_key = key
                        break

        sample_ml = batch_count * vessels * vol_per_vessel
        total_ml  = sample_ml + std_once

        return {
            "medium_name": conds.get("medium_name", ""),
            "apparatus": conds.get("apparatus", ""),
            "speed_rpm": conds.get("speed_rpm"),
            "sampling_time": conds.get("sampling_time", ""),
            "volume_per_vessel_ml": vol_per_vessel,
            "vessels_per_batch": vessels,
            "batch_count": batch_count,
            "sample_medium_ml": round(sample_ml, 1),
            "standard_medium_ml_once": round(std_once, 1),
            "total_medium_ml": round(total_ml, 1),
            "_std_medium_key": matched_std_key,
        }
    return None


def merge_all_results(
    results: list[dict],
    strength_configs: list[dict],  # [{strength, test_items: [{name, batch_count}]}]
) -> dict:
    """(함량 × 시험항목) 계산 결과 전체를 합산합니다."""
    first = results[0]

    if len(results) == 1:
        r = dict(results[0])
        r["strength_configs"] = strength_configs
        return r

    # Solutions: solution_name 키로 합산. 표준액·표준원액은 함량·배치 무관 1회만.
    sol_map: dict[str, dict] = {}
    for r in results:
        for sol in r["solutions"]:
            key = sol.get("solution_name", "")
            if key not in sol_map:
                sol_map[key] = dict(sol)
            elif not _RE_STD_PREP.search(key):
                a = sol_map[key].get("theoretical_volume_ml") or 0
                b = sol.get("theoretical_volume_ml") or 0
                if _RE_MP.match(key):
                    # 이동상: 고정 주입은 1회만, 검액(배치비례) 주입만 합산
                    fixed_raw = sol_map[key].get("_hplc_fixed_raw") or 0
                    sum_batch = (sol_map[key].get("_hplc_batch_raw") or 0) + (sol.get("_hplc_batch_raw") or 0)
                    if fixed_raw:
                        # 균일성끼리 합산 시: 브라켓팅 수를 총 검액 수로 재계산
                        nbf_raw_a = sol_map[key].get("_hplc_non_bracket_fixed_raw")
                        nbf_raw_b = sol.get("_hplc_non_bracket_fixed_raw")
                        if nbf_raw_a is not None and nbf_raw_b is not None:
                            total_samp = (sol_map[key].get("_hplc_total_samples") or 0) + (sol.get("_hplc_total_samples") or 0)
                            flow_rtime = sol_map[key].get("_hplc_flow_rtime") or sol.get("_hplc_flow_rtime") or 0
                            interval   = sol_map[key].get("_hplc_interval") or sol.get("_hplc_interval") or 20
                            n_brk = math.ceil(total_samp / interval) + 1 if total_samp else 1
                            new_fixed_raw = nbf_raw_a + flow_rtime * n_brk
                            combined_raw = new_fixed_raw + sum_batch
                            sol_map[key]["_hplc_total_samples"] = total_samp
                            sol_map[key]["_hplc_fixed_raw"] = new_fixed_raw
                            sol_map[key]["_hplc_interval"] = interval
                        else:
                            combined_raw = fixed_raw + sum_batch
                        sol_map[key]["theoretical_volume_ml"] = math.ceil(combined_raw / 50) * 50
                        sol_map[key]["_hplc_batch_raw"] = sum_batch
                    else:
                        sol_map[key]["theoretical_volume_ml"] = round(a + b, 1)
                elif _RE_BUF.match(key):
                    # 완충액: mp 최종 계산 후 비율로 재계산 — 지금은 패스
                    pass
                else:
                    sol_map[key]["theoretical_volume_ml"] = round(a + b, 1)
                # vol_per_batch_ml이 None이면 후속 항목 값 사용
                if sol_map[key].get("volume_per_batch_ml") is None and sol.get("volume_per_batch_ml"):
                    sol_map[key]["volume_per_batch_ml"] = sol["volume_per_batch_ml"]
            # else: 표준액·표준원액 중복 → 스킵

    # 완충액: 동일 이동상의 최종 이론량 × 완충액 비율로 재계산
    for buf_key, buf_sol in sol_map.items():
        if not _RE_BUF.match(buf_key):
            continue
        ratio = buf_sol.get("_hplc_buf_ratio")
        if not ratio:
            continue
        for mp_key, mp_sol in sol_map.items():
            if _RE_MP.match(mp_key) and mp_sol.get("_hplc_fixed_raw"):
                buf_sol["theoretical_volume_ml"] = round(mp_sol["theoretical_volume_ml"] * ratio, 1)
                break

    # Glassware: (type, size, source_prep, strength, test_item) 키로 합산.
    # 표준액·표준원액 및 피펫/메스실린더는 함량 무관 1회만 계산.
    gw_map: dict[tuple, dict] = {}
    for r in results:
        for gw in r["glassware"]:
            src = gw.get("source_prep", "")
            is_std    = bool(_RE_STD_PREP.search(src))
            is_pipette = gw.get("type", "") in ("pipette", "graduated cylinder")
            str_key      = "" if is_std else gw.get("strength", "")
            # 표준액 초자는 어느 시험항목에서 나왔든 1회만 집계 (Rule 17)
            test_item_key = "" if is_std else gw.get("test_item", "")
            key = (gw["type"], gw["size"], src, str_key, test_item_key)
            if key not in gw_map:
                entry = dict(gw)
                entry["strength"] = str_key
                gw_map[key] = entry
            elif not is_std:
                # 피펫은 count_per_batch 누적 없이 total_count만 합산 (각 배치마다 별도 피펫 필요)
                if not is_pipette:
                    gw_map[key]["count_per_batch"] += gw["count_per_batch"]
                gw_map[key]["total_count"] += gw["total_count"]
            # else: 표준액·표준원액 중복 → 스킵

    # Filters: (size_um, material, manufacturer, filter_type, source_prep, test_item) 키로 합산
    fi_map: dict[tuple, dict] = {}
    for r in results:
        for fi in r["filters"]:
            key = (fi.get("size_um"), fi.get("material"), fi.get("manufacturer"), fi.get("filter_type"), fi.get("source_prep", ""), fi.get("test_item", ""))
            if key not in fi_map:
                fi_map[key] = dict(fi)
            else:
                fi_map[key]["count_per_batch"] += fi["count_per_batch"]
                fi_map[key]["total_count"] += fi["total_count"]

    # Pipettes: 중복 제거
    pip_map: dict[float, dict] = {}
    for r in results:
        for pip in r["pipettes"]:
            vol = pip["volume_ml"]
            if vol not in pip_map:
                pip_map[vol] = dict(pip)

    # Dissolution medium: Method A/B는 동일 vessel 사용 → sample은 max
    # Rule 17: 표준액 용출량은 동일 std_medium_key 공유 시 1회만 집계
    dm_list = [r["dissolution_medium"] for r in results if r.get("dissolution_medium")]
    merged_dm = None
    if dm_list:
        merged_dm = dict(dm_list[0])
        merged_dm["sample_medium_ml"] = max(d["sample_medium_ml"] for d in dm_list)
        seen_std_keys: set[str] = set()
        std_total = 0.0
        for d in dm_list:
            k = d.get("_std_medium_key", "")
            if k and k in seen_std_keys:
                continue  # 같은 표준액 조제 키 → 중복 제외
            if k:
                seen_std_keys.add(k)
            std_total += d["standard_medium_ml_once"]
        merged_dm["standard_medium_ml_once"] = round(std_total, 1)
        merged_dm["total_medium_ml"] = round(
            merged_dm["sample_medium_ml"] + merged_dm["standard_medium_ml_once"], 1
        )

    # Standard names 중복 제거
    rs_seen: set[str] = set()
    rs_names: list[str] = []
    for r in results:
        for nm in r.get("standard_names", []):
            if nm.lower() not in rs_seen:
                rs_seen.add(nm.lower())
                rs_names.append(nm)

    return {
        "product_name": first["product_name"],
        "doc_no": first["doc_no"],
        "stm_file": first["stm_file"],
        "strength_configs": strength_configs,
        "solutions": list(sol_map.values()),
        "glassware": list(gw_map.values()),
        "filters": list(fi_map.values()),
        "pipettes": sorted(pip_map.values(), key=lambda x: x["volume_ml"]),
        "dissolution_medium": merged_dm,
        "standard_names": rs_names,
    }


def scale_ingredients(
    ingredients: list[dict],
    volume_per_batch_ml: float,
    actual_prep_ml: float,
) -> list[dict]:
    if not volume_per_batch_ml or volume_per_batch_ml == 0:
        return ingredients
    scale = actual_prep_ml / volume_per_batch_ml
    result = []
    for ing in ingredients:
        scaled_amount = round(ing["amount"] * scale, 4)
        result.append({**ing, "scaled_amount": scaled_amount})
    return result
