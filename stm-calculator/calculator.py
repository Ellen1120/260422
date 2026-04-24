"""
배치 수 기반 시험 자원 계산 엔진
"""
from __future__ import annotations

import math
import re

_RE_MP    = re.compile(r'^mobile\s*phase', re.IGNORECASE)
_RE_BUF   = re.compile(r'^buffer\b', re.IGNORECASE)
_RE_SAMPLE_SOL  = re.compile(r'sample', re.IGNORECASE)
_RE_UNIFORMITY  = re.compile(r'uniformity', re.IGNORECASE)
_RE_ASSAY       = re.compile(r'^assay$', re.IGNORECASE)
_RE_RATIO = re.compile(
    r'Mix\s+(.+?)\s+and\s+(.+?)\s+in\s+the\s+ratio\s+of\s+'
    r'(\d+(?:\.\d+)?)\s*:\s*(\d+(?:\.\d+)?)',
    re.IGNORECASE,
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
):
    """단일 시험항목의 preparations를 순회하며 solutions/glassware/filter 집계."""
    sample_count = _sample_count_for_item(item_name)

    for prep in preparations:
        sol_name = prep.get("solution_name", prep.get("section_name", ""))
        is_sample_prep = bool(_RE_SAMPLE_SOL.search(sol_name))
        if skip_sample and is_sample_prep:
            continue

        vol = prep.get("volume_per_batch_ml")
        theoretical = round(vol * batch_count, 1) if vol else None

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

        # 초자 집계
        for gw in prep.get("glassware", []):
            key = (gw.get("type", ""), gw.get("size", ""))
            per_batch = gw.get("count_per_batch", 1) * per_sample
            if key not in glassware_agg:
                glassware_agg[key] = {
                    "type": gw.get("type", ""),
                    "size": gw.get("size", ""),
                    "count_per_batch": per_batch,
                    "total_count": per_batch * batch_count,
                }
            else:
                glassware_agg[key]["count_per_batch"] += per_batch
                glassware_agg[key]["total_count"] += per_batch * batch_count

        # 필터 집계
        for fi in prep.get("filters", []):
            mat  = fi.get("material", "")
            mfr  = fi.get("manufacturer", "")
            ftype = fi.get("filter_type", "syringe")
            size  = fi.get("size_um", 0.45)
            key = (size, mat, mfr, ftype)
            per_batch = 1 * per_sample
            if key not in filter_agg:
                filter_agg[key] = {
                    "size_um": size,
                    "material": mat,
                    "manufacturer": mfr,
                    "filter_type": ftype,
                    "count_per_batch": per_batch,
                    "total_count": per_batch * batch_count,
                }
            else:
                filter_agg[key]["count_per_batch"] += per_batch
                filter_agg[key]["total_count"] += per_batch * batch_count


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
                skip_sample=True,  # sample solution은 제외 (Uniformity 고유 것 사용)
            )

        _process_preparations(
            item_name=item["name"],
            preparations=item.get("preparations", []),
            batch_count=batch_count,
            solutions=solutions,
            glassware_agg=glassware_agg,
            filter_agg=filter_agg,
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
            m = _RE_RATIO.search(sol.get("preparation_text", ""))
            if m:
                sub_a, sub_b = m.group(1).strip(), m.group(2).strip()
                r_a, r_b = float(m.group(3)), float(m.group(4))
                t = r_a + r_b
                if t > 0:
                    sol["ingredients"] = [
                        {"name": sub_a, "amount": round(r_a / t * mp_vol, 1), "unit": "mL"},
                        {"name": sub_b, "amount": round(r_b / t * mp_vol, 1), "unit": "mL"},
                    ]
                    if "buffer" in sub_a.lower():
                        total_buffer_needed += mp_vol * (r_a / t)
                    elif "buffer" in sub_b.lower():
                        total_buffer_needed += mp_vol * (r_b / t)

        if total_buffer_needed > 0 and buf_sols:
            for sol in buf_sols:
                sol["theoretical_volume_ml"] = round(total_buffer_needed, 1)

    dissolution_medium = _calc_dissolution_medium(
        product, strength, test_item_names, batch_count
    )

    # 피펫/메스 실린더: solutions의 모든 재료(mL)에서 고유 규격 추출, 공용 → count=1
    _RE_BTV = re.compile(
        r'(?:dilute to volume|bring to volume|make up to volume|make to volume)\s+with\s+([^.,\n]+)',
        re.IGNORECASE,
    )
    pipette_agg: dict[float, dict] = {}
    for sol in solutions:
        prep_text = sol.get("preparation_text", "")
        btv_m = _RE_BTV.search(prep_text)
        btv_solvent = btv_m.group(1).strip().lower() if btv_m else ""

        for ing in sol.get("ingredients", []):
            if ing.get("unit", "").strip().lower() != "ml":
                continue
            vol = round(float(ing.get("amount", 0)), 1)
            if vol <= 0:
                continue
            # 표선 용매와 동일한 재료는 메스 실린더로 계량할 필요 없음
            ing_name = ing.get("name", "").strip().lower()
            if btv_solvent and (ing_name in btv_solvent or btv_solvent.startswith(ing_name)):
                continue
            if vol not in pipette_agg:
                pipette_agg[vol] = {
                    "volume_ml": vol,
                    "type": "홀 피펫" if vol <= 25 else "메스 실린더",
                    "count": 1,
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
        "pipettes": sorted(pipette_agg.values(), key=lambda x: x["volume_ml"]),
        "dissolution_medium": dissolution_medium,
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
