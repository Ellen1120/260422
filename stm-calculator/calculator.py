"""
배치 수 기반 시험 자원 계산 엔진

흐름:
1. 배치 수 × 1배치 기준량 = 이론 조제량 산출
2. 시험자가 실제 조제량을 결정
3. scale_factor = 실제 조제량 / 1배치 기준량 → 시약 필요량 산출
"""
from __future__ import annotations

import re


def calculate_resources(
    product: dict,
    strength: str,
    test_item_names: list[str],
    batch_count: int,
) -> dict:
    """
    선택된 시험항목들의 용액·초자 목록 반환.
    이론 조제량 = batch_count × volume_per_batch_ml (여유분 없음).
    """
    solutions: list[dict] = []
    glassware_agg: dict[tuple, dict] = {}

    for item in product.get("test_items", []):
        if item["name"] not in test_item_names:
            continue

        for prep in item.get("preparations", []):
            vol = prep.get("volume_per_batch_ml")
            theoretical = round(vol * batch_count, 1) if vol else None

            solutions.append(
                {
                    "test_item": item["name"],
                    "solution_name": prep.get("solution_name", prep.get("section_name", "")),
                    "section_name": prep.get("section_name", ""),
                    "volume_per_batch_ml": vol,
                    "theoretical_volume_ml": theoretical,
                    "preparation_text": prep.get("preparation_text", ""),
                    "ingredients": prep.get("ingredients", []),
                }
            )

            for gw in prep.get("glassware", []):
                key = (gw.get("type", ""), gw.get("size", ""))
                per_batch = gw.get("count_per_batch", 1)
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

    # 용출 시험 선택 시 시험액 필요량 계산
    dissolution_medium = _calc_dissolution_medium(
        product, strength, test_item_names, batch_count
    )

    return {
        "product_name": product.get("product_name", ""),
        "doc_no": product.get("doc_no"),
        "stm_file": product.get("stm_file", ""),
        "strength": strength,
        "test_items": test_item_names,
        "batch_count": batch_count,
        "solutions": solutions,
        "glassware": list(glassware_agg.values()),
        "dissolution_medium": dissolution_medium,
    }


def _calc_dissolution_medium(
    product: dict,
    strength: str,
    test_item_names: list[str],
    batch_count: int,
) -> dict | None:
    """
    용출 시험이 선택된 경우 시험액 총 필요량 계산.
    공식: (배치 수 × 검체 수 × 용출 부피) + 표준액용 시험액(1회)
    """
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

        vessels = conds.get("vessels_per_batch", 6)
        std_by_str = conds.get("standard_medium_ml_by_strength", {})
        std_once = float(std_by_str.get(strength, 0)) if std_by_str else 0.0

        sample_ml = batch_count * vessels * vol_per_vessel
        total_ml = sample_ml + std_once

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
    """
    실제 조제량 기준으로 시약 필요량 산출.
    scale_factor = actual_prep_ml / volume_per_batch_ml
    """
    if not volume_per_batch_ml or volume_per_batch_ml == 0:
        return ingredients
    scale = actual_prep_ml / volume_per_batch_ml
    result = []
    for ing in ingredients:
        scaled_amount = round(ing["amount"] * scale, 4)
        result.append({**ing, "scaled_amount": scaled_amount})
    return result
