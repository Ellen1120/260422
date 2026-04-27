"""
배치 수 기반 시험 자원 계산 엔진
"""
from __future__ import annotations

import math
import re

_RE_MP    = re.compile(r'^mobile\s*phase|^이동상', re.IGNORECASE)
_RE_BUF   = re.compile(r'^buffer\b|^완충액', re.IGNORECASE)
_RE_DILUENT = re.compile(r'^diluent|^희석액', re.IGNORECASE)
_RE_SAMPLE_SOL  = re.compile(r'sample', re.IGNORECASE)
_RE_UNIFORMITY  = re.compile(r'uniformity', re.IGNORECASE)
_RE_ASSAY       = re.compile(r'^assay$', re.IGNORECASE)
# 표준액·표준원액: 배치 수·함량 수 무관 1회만 조제
_RE_STD_PREP    = re.compile(r'standard\s+(?:stock\s+)?solution', re.IGNORECASE)
# 함량 지정 조제 패턴: "Standard solution (for 50 mg)"
_RE_FOR_STRENGTH = re.compile(r'\(\s*for\s+(.+?)\s*\)', re.IGNORECASE)
# 피펫 집계 대상: 표준원액/표준액/이동상/희석액 포함 (검액 제외 - 대용량 vessel 전달)
_RE_PIPETTE_SOURCE  = re.compile(r'sample|standard|이동상|희석액|표준원액|표준액|mobile\s*phase|diluent', re.IGNORECASE)
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
            per_batch = gw.get("count_per_batch", 1) * per_sample
            if key not in glassware_agg:
                glassware_agg[key] = {
                    "type": gw.get("type", ""),
                    "size": gw.get("size", ""),
                    "source_prep": sol_name,
                    "strength": gw_strength,
                    "test_item": gw_test_item,
                    "count_per_batch": per_batch,
                    "total_count": per_batch * effective_batches,
                }
            else:
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
            sol["volume_per_batch_ml"]   = mp_vol
            sol["theoretical_volume_ml"] = mp_vol
            prep_text = sol.get("preparation_text", "")
            m = _RE_RATIO.search(prep_text) or _RE_RATIO_KO.search(prep_text)
            if m:
                sub_a, sub_b = m.group(1).strip(), m.group(2).strip()
                r_a, r_b = float(m.group(3)), float(m.group(4))
                t = r_a + r_b
                if t > 0:
                    sol["ingredients"] = [
                        {"name": sub_a, "amount": round(r_a / t * mp_vol, 1), "unit": "mL"},
                        {"name": sub_b, "amount": round(r_b / t * mp_vol, 1), "unit": "mL"},
                    ]
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
                if (ing.get("name") or "") == sol_name_exact and \
                   ing.get("unit", "").lower() == "ml":
                    total_used += float(ing.get("amount", 0)) * scale
        if total_used > 0:
            prep_text = sol.get("preparation_text", "")
            m = _RE_RATIO.search(prep_text) or _RE_RATIO_KO.search(prep_text)
            sol["theoretical_volume_ml"] = round(total_used, 1)
            sol["volume_per_batch_ml"]   = total_used
            if m:
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
            nm = (std.get("std_name") or "").strip()
            if nm and nm.lower() not in rs_seen:
                rs_seen.add(nm.lower())
                standard_names.append(nm)

    dissolution_medium = _calc_dissolution_medium(
        product, strength, test_item_names, batch_count
    )

    # 피펫/메스 실린더: mL 재료 규격 추출 → 초자 목록(glassware_agg)에 통합
    _RE_BTV = re.compile(
        r'(?:dilute to volume|bring to volume|make up to volume|make to volume)\s+with\s+([^.,\n]+)',
        re.IGNORECASE,
    )
    for sol in solutions:
        if not _RE_PIPETTE_SOURCE.search(sol.get("solution_name", "")):
            continue

        prep_text = sol.get("preparation_text", "")
        btv_m = _RE_BTV.search(prep_text)
        btv_solvent = btv_m.group(1).strip().lower() if btv_m else ""
        vol_per_batch = sol.get("volume_per_batch_ml") or 0

        for ing in sol.get("ingredients", []):
            if ing.get("unit", "").strip().lower() != "ml":
                continue
            vol = round(float(ing.get("amount", 0)), 1)
            if vol <= 0:
                continue
            ing_name = ing.get("name", "").strip().lower()
            # 표선 용매 자체이고 총량의 50% 이상이면 BTV → 피펫 불필요
            if btv_solvent and (ing_name in btv_solvent or btv_solvent.startswith(ing_name)):
                if vol_per_batch == 0 or vol >= vol_per_batch * 0.5:
                    continue
            gtype = "홀 피펫" if vol <= 25 else "메스 실린더"
            size  = f"{vol} mL"
            key   = (gtype, size, gtype, "", "")
            if key not in glassware_agg:
                glassware_agg[key] = {
                    "type": gtype,
                    "size": size,
                    "source_prep": gtype,
                    "strength": "",
                    "test_item": "",
                    "count_per_batch": 1,
                    "total_count": 1,
                }

    return {
        "product_name": product.get("product_name", ""),
        "doc_no": product.get("doc_no"),
        "stm_file": product.get("stm_file", ""),
        "strength": strength,
        "test_items": test_item_names,
        "batch_count": batch_count,
        "solutions": solutions,
        "glassware": list(glassware_agg.values()),
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
        if not re.match(r"^dissolution\s*$", item["name"], re.IGNORECASE):
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
            is_pipette = gw.get("type", "") in ("홀 피펫", "메스 실린더")
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

    # Dissolution medium: 합산
    dm_list = [r["dissolution_medium"] for r in results if r.get("dissolution_medium")]
    merged_dm = None
    if dm_list:
        merged_dm = dict(dm_list[0])
        merged_dm["sample_medium_ml"] = round(sum(d["sample_medium_ml"] for d in dm_list), 1)
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
