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
_RE_SAMPLE_SOL  = re.compile(r'sample', re.IGNORECASE)
_RE_UNIFORMITY  = re.compile(r'uniformity', re.IGNORECASE)
_RE_ASSAY       = re.compile(r'^assay$', re.IGNORECASE)
# 표준액·표준원액: 배치 수·함량 수 무관 1회만 조제
_RE_STD_PREP    = re.compile(r'standard\s+(?:stock\s+)?solution', re.IGNORECASE)
# 함량 지정 조제 패턴: "Standard solution (for 50 mg)"
_RE_FOR_STRENGTH = re.compile(r'\(\s*for\s+(.+?)\s*\)', re.IGNORECASE)
# 피펫 집계 대상: 검액 및 표준액만 포함 (용액 조제 시약류 제외 규칙 반영)
_RE_PIPETTE_SOURCE  = re.compile(r'sample|standard|표준원액|표준액', re.IGNORECASE)

# Rule 3: 용액 조제 목록에서 제외할 시험 용액 (표준액, 표준원액, 검액)
_RE_TEST_SOLUTION = re.compile(
    r'\bstandard\b.*\bsolution\b|\bsample\b|표준원액|표준액|검액',
    re.IGNORECASE,
)
# Rule 1: 초자 목록에 포함할 소스 준비 (standard/sample 준비에서 나온 초자만 유지)
_RE_GW_KEEP_SOURCE = re.compile(r'standard|sample|표준|검액', re.IGNORECASE)

_RE_SAMPLE_OR_STD   = re.compile(r'sample|standard', re.IGNORECASE)
_RE_RATIO = re.compile(
    r'Mix\s+(.+?)\s+and\s+(.+?)\s+in\s+the\s+ratio\s+of\s+'
    r'(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)',
    re.IGNORECASE,
)
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
    if re.search(r'dissolution', name):
        return 6
    if re.search(r'uniformity', name):
        return 10
    return 1


def _hplc_mp_volume(hplc: dict, batch_count: int) -> float | None:
    flow  = hplc.get("flow_rate_ml_min")
    rtime = hplc.get("run_time_min")
    if not flow or not rtime:
        return None
    injs = hplc.get("injections", [])
    total_inj = sum(
        inj["count"] * batch_count if inj.get("scales_with_batch") else inj["count"]
        for inj in injs
    ) if injs else 1
    raw = flow * (rtime * total_inj + 30) * 1.5
    return math.ceil(raw / 50) * 50


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
        if strength:
            m_str = _RE_FOR_STRENGTH.search(sol_name)
            if m_str and m_str.group(1).strip().lower() != strength.strip().lower():
                continue

        is_sample_prep = bool(_RE_SAMPLE_SOL.search(sol_name))
        if skip_sample and is_sample_prep:
            continue

        # 표준액·표준원액은 배치 수에 무관하게 1회만 조제
        is_std_prep      = bool(_RE_STD_PREP.search(sol_name))
        effective_batches = 1 if is_std_prep else batch_count

        vol = prep.get("volume_per_batch_ml")
        theoretical = round(vol * effective_batches, 1) if vol else None

        solutions.append({
            "test_item": item_name,
            "solution_name": sol_name,
            "section_name": prep.get("section_name", ""),
            "volume_per_batch_ml": vol,
            "theoretical_volume_ml": theoretical,
            "preparation_text": prep.get("preparation_text", ""),
            "ingredients": list(prep.get("ingredients", [])),
        })

        per_sample = sample_count if is_sample_prep else 1

        # 초자 집계 — 조제별/함량별/시험항목별 별도 행 (검액·표준액만 strength·test_item 포함)
        is_sample_std = bool(_RE_SAMPLE_OR_STD.search(sol_name))
        gw_strength  = (strength or "") if is_sample_std else ""
        gw_test_item = item_name        if is_sample_std else ""
        for gw in prep.get("glassware", []):
            key = (gw.get("type", ""), gw.get("size", ""), sol_name, gw_strength, gw_test_item)
            # mortar는 공유 장비이므로 배치 수·검액 수 무관하게 1개
            is_shared = gw.get("type", "") == "mortar"
            per_batch = 1 if is_shared else gw.get("count_per_batch", 1) * per_sample
            if key not in glassware_agg:
                glassware_agg[key] = {
                    "type": gw.get("type", ""),
                    "size": gw.get("size", ""),
                    "source_prep": sol_name,
                    "strength": gw_strength,
                    "test_item": gw_test_item,
                    "count_per_batch": per_batch,
                    "total_count": 1 if is_shared else per_batch * effective_batches,
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

    # Uniformity가 선택된 경우 Assay의 준비 항목을 함께 가져올 Assay 아이템 탐색
    assay_item_for_uniformity: dict | None = None
    has_uniformity = any(_RE_UNIFORMITY.search(n) for n in test_item_names)
    if has_uniformity:
        for it in product.get("test_items", []):
            if _RE_ASSAY.match(it["name"]):
                assay_item_for_uniformity = it
                break

    for item in product.get("test_items", []):
        if item["name"] not in test_item_names:
            continue

        if _RE_UNIFORMITY.search(item["name"]) and assay_item_for_uniformity:
            # Uniformity: Assay의 비-sample 조제를 먼저 포함 (buffer, mobile phase, diluent 등)
            _process_preparations(
                item_name=item["name"],
                preparations=assay_item_for_uniformity.get("preparations", []),
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

        # Uniformity → Assay hplc_conditions 사용
        if _RE_UNIFORMITY.search(item["name"]) and assay_item_for_uniformity:
            hplc = assay_item_for_uniformity.get("hplc_conditions")
            item_name_for_hplc = item["name"]
        else:
            if item["name"] not in test_item_names:
                continue
            hplc = item.get("hplc_conditions")

        if not hplc:
            continue
        if item["name"] not in test_item_names:
            continue

        mp_vol = _hplc_mp_volume(hplc, batch_count)
        if not mp_vol:
            continue

        mp_sols  = [s for s in solutions if s["test_item"] == item_name_for_hplc and _RE_MP.match(s["solution_name"])]
        buf_sols = [s for s in solutions if s["test_item"] == item_name_for_hplc and _RE_BUF.match(s["solution_name"])]

        total_buffer_needed = 0.0
        for sol in mp_sols:
            # 기존 조제 부피 (스케일링 전) 백업
            old_vol = sol.get("volume_per_batch_ml") or 1000.0
            sol["volume_per_batch_ml"]   = mp_vol
            sol["theoretical_volume_ml"] = mp_vol
            prep_text = sol.get("preparation_text", "")
            m3 = _RE_RATIO_KO3.search(prep_text)
            if m3:
                sub_a, sub_b, sub_c = m3.group(1).strip(), m3.group(2).strip(), m3.group(3).strip()
                r_a, r_b, r_c = float(m3.group(4)), float(m3.group(5)), float(m3.group(6))
                t = r_a + r_b + r_c
                if t > 0:
                    new_ings = [
                        {"name": sub_a, "amount": round(r_a / t * mp_vol, 1), "unit": "mL"},
                        {"name": sub_b, "amount": round(r_b / t * mp_vol, 1), "unit": "mL"},
                        {"name": sub_c, "amount": round(r_c / t * mp_vol, 1), "unit": "mL"},
                    ]
                    # 기존 고체 성분(g, mg) 보존 및 스케일링
                    scale = mp_vol / old_vol
                    for old_ing in sol.get("ingredients", []):
                        if old_ing.get("unit", "").lower() in ["g", "mg"]:
                            scaled_ing = dict(old_ing)
                            scaled_ing["amount"] = round(float(old_ing["amount"]) * scale, 2)
                            new_ings.append(scaled_ing)
                    sol["ingredients"] = new_ings
                    for sub, r in [(sub_a, r_a), (sub_b, r_b), (sub_c, r_c)]:
                        if "buffer" in sub.lower() or "완충액" in sub:
                            total_buffer_needed += mp_vol * (r / t)
                            break
            else:
                m = _RE_RATIO.search(prep_text) or _RE_RATIO_KO.search(prep_text)
                if not m:
                    m = _RE_RATIO_KO2_MIT.search(prep_text) or _RE_RATIO_KO2_WA.search(prep_text)
                if m:
                    sub_a, sub_b = m.group(1).strip(), m.group(2).strip()
                    r_a, r_b = float(m.group(3)), float(m.group(4))
                    t = r_a + r_b
                    if t > 0:
                        new_ings = [
                            {"name": sub_a, "amount": round(r_a / t * mp_vol, 1), "unit": "mL"},
                            {"name": sub_b, "amount": round(r_b / t * mp_vol, 1), "unit": "mL"},
                        ]
                        # 기존 고체 성분(g, mg) 보존 및 스케일링
                        scale = mp_vol / old_vol
                        for old_ing in sol.get("ingredients", []):
                            if old_ing.get("unit", "").lower() in ["g", "mg"]:
                                scaled_ing = dict(old_ing)
                                scaled_ing["amount"] = round(float(old_ing["amount"]) * scale, 2)
                                new_ings.append(scaled_ing)
                        sol["ingredients"] = new_ings
                        
                        if "buffer" in sub_a.lower() or "완충액" in sub_a:
                            total_buffer_needed += mp_vol * (r_a / t)
                        elif "buffer" in sub_b.lower() or "완충액" in sub_b:
                            total_buffer_needed += mp_vol * (r_b / t)

        if total_buffer_needed > 0 and buf_sols:
            for sol in buf_sols:
                sol["theoretical_volume_ml"] = round(total_buffer_needed, 1)

    # 희석액/diluent: 다른 조제에서 사용된 총량 합산 → 이론량 + 비율 재료 업데이트
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
            scale = (other_theor / other_vol) if other_vol else 1.0
            for ing in other.get("ingredients", []):
                if (ing.get("name") or "").lower() == sol_name_exact.lower() and \
                   ing.get("unit", "").lower() == "ml":
                    total_used += float(ing.get("amount", 0)) * scale
        if total_used > 0:
            prep_text = sol.get("preparation_text", "")
            sol["theoretical_volume_ml"] = round(total_used, 1)
            sol["volume_per_batch_ml"]   = total_used
            m3 = _RE_RATIO_KO3.search(prep_text)
            m = None if m3 else (_RE_RATIO.search(prep_text) or _RE_RATIO_KO.search(prep_text) or _RE_RATIO_KO2_MIT.search(prep_text) or _RE_RATIO_KO2_WA.search(prep_text))
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

    # 표준품 이름 수집 (STD Name 표 기반)
    rs_seen: set[str] = set()
    standard_names: list[str] = []
    for item in product.get("test_items", []):
        if item["name"] not in test_item_names:
            continue
        item_stds = item.get("standards", [])
        # Uniformity: 자체 standards 없으면 Assay 표준품 사용
        if not item_stds and _RE_UNIFORMITY.search(item["name"]) and assay_item_for_uniformity:
            item_stds = assay_item_for_uniformity.get("standards", [])
        for std in item_stds:
            nm = _translate_std_name((std.get("std_name") or "").strip())
            if nm and nm.lower() not in rs_seen:
                rs_seen.add(nm.lower())
                standard_names.append(nm)

    dissolution_medium = _calc_dissolution_medium(
        product, strength, test_item_names, batch_count
    )

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
                _RE_RATIO_KO2_WA.search(prep_text)):
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
            key  = ("pipette", size, sol_name, pip_strength, pip_test_item)
            if key not in glassware_agg:
                glassware_agg[key] = {
                    "type": "pipette",
                    "size": size,
                    "source_prep": sol_name,
                    "strength": pip_strength,
                    "test_item": pip_test_item,
                    "count_per_batch": 1,
                    "total_count": 1,
                }

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
    glassware_list = [
        g for g in glassware_agg.values()
        if _RE_GW_KEEP_SOURCE.search(g.get("source_prep", ""))
    ]

    # Rule 7: 용출 시험 시 900mL 피펫은 불필요 (용출 배지를 피펫으로 옮기지 않음)
    glassware_list = [
        g for g in glassware_list
        if not (g.get("type") == "pipette" and "900" in str(g.get("size", "")))
    ]

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
        std_once = float(std_by_str.get(strength, 0)) if std_by_str else 0.0

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
                sol_map[key]["theoretical_volume_ml"] = round(a + b, 1)
                # vol_per_batch_ml이 None이면 후속 항목 값 사용
                if sol_map[key].get("volume_per_batch_ml") is None and sol.get("volume_per_batch_ml"):
                    sol_map[key]["volume_per_batch_ml"] = sol["volume_per_batch_ml"]
            # else: 표준액·표준원액 중복 → 스킵

    # Glassware: (type, size, source_prep, strength, test_item) 키로 합산.
    # 표준액·표준원액 및 피펫/메스실린더는 함량 무관 1회만 계산.
    gw_map: dict[tuple, dict] = {}
    for r in results:
        for gw in r["glassware"]:
            src = gw.get("source_prep", "")
            is_std    = bool(_RE_STD_PREP.search(src))
            is_pipette = gw.get("type", "") in ("pipette", "graduated cylinder")
            str_key = "" if is_std else gw.get("strength", "")
            key = (gw["type"], gw["size"], src, str_key, gw.get("test_item", ""))
            if key not in gw_map:
                entry = dict(gw)
                entry["strength"] = str_key
                gw_map[key] = entry
            elif not is_std and not is_pipette:
                gw_map[key]["count_per_batch"] += gw["count_per_batch"]
                gw_map[key]["total_count"] += gw["total_count"]
            # else: 표준액·표준원액·피펫/메스실린더 중복 → 스킵

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

    # Dissolution medium: Method A/B는 동일 vessel 사용 → sample은 max, standard는 합산
    dm_list = [r["dissolution_medium"] for r in results if r.get("dissolution_medium")]
    merged_dm = None
    if dm_list:
        merged_dm = dict(dm_list[0])
        merged_dm["sample_medium_ml"] = max(d["sample_medium_ml"] for d in dm_list)
        merged_dm["standard_medium_ml_once"] = round(sum(d["standard_medium_ml_once"] for d in dm_list), 1)
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
