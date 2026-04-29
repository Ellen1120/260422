"""
STM 자원 계산기 - FastAPI 서버
"""
from __future__ import annotations

import json
import os
import re
import threading
from pathlib import Path

from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

from parser import build_knowledge_base, load_knowledge_base, save_knowledge_base
from calculator import calculate_resources, scale_ingredients, merge_all_results
from standards_db import lookup as lookup_standards
from column_db import lookup as lookup_columns

CONFIG_PATH = Path(__file__).parent / "data" / "config.json"


def _load_config() -> dict:
    if CONFIG_PATH.exists():
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    return {}


def _save_config(cfg: dict) -> None:
    CONFIG_PATH.parent.mkdir(exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def _get_api_key() -> str | None:
    """환경변수 → config.json → Claude Code 자격증명 순으로 API 키 탐색."""
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    if key:
        return key
    cfg = _load_config()
    if cfg.get("api_key"):
        return cfg["api_key"]
    # Claude Code가 저장한 OAuth 토큰 자동 사용
    cred_path = Path.home() / ".claude" / ".credentials.json"
    if cred_path.exists():
        try:
            cred = json.loads(cred_path.read_text(encoding="utf-8"))
            token = cred.get("claudeAiOauth", {}).get("accessToken", "")
            if token:
                return token
        except Exception:
            pass
    return None

app = FastAPI(title="STM 시험 자원 계산기")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── 앱 상태 ──────────────────────────────────────────────
_state = {
    "products": [],
    "parsing": False,
    "parse_error": None,
    "parse_log": [],
}
_state_lock = threading.Lock()


def _needs_reparse() -> bool:
    """STM 폴더에 knowledge_base.json보다 새로운 .docx 파일이 있으면 True."""
    from parser import STM_FOLDER, KB_PATH
    if not KB_PATH.exists():
        return True
    kb_mtime = KB_PATH.stat().st_mtime
    for f in STM_FOLDER.glob("*.docx"):
        if not f.name.startswith("~$") and f.stat().st_mtime > kb_mtime:
            return True
    return False


@app.on_event("startup")
async def startup():
    _state["products"] = load_knowledge_base()
    print(f"Loaded {len(_state['products'])} products from knowledge base.")


def _auto_reparse():
    def _log(msg: str):
        print(msg)
        _state["parse_log"].append(msg)
    with _state_lock:
        _state["parsing"] = True
        _state["parse_error"] = None
        _state["parse_log"] = []
    try:
        products = build_knowledge_base(progress_callback=_log)
        with _state_lock:
            _state["products"] = products
        _log(f"완료: {len(products)}개 제품 파싱됨")
    except Exception as e:
        with _state_lock:
            _state["parse_error"] = str(e)
        _log(f"오류: {e}")
    finally:
        with _state_lock:
            _state["parsing"] = False


# ── 제품 목록 ─────────────────────────────────────────────
def _extract_code_no(stm_file: str) -> str:
    """'STM-300599-T4.doc.docx' → '300599'  ([접두어]-[코드번호]-[접미어] 중 가운데)"""
    parts = stm_file.split("-")
    return parts[1] if len(parts) >= 2 else stm_file


@app.get("/api/products")
def get_products():
    return [
        {
            "id": p["id"],
            "name": p["product_name"],
            "stm_file": p["stm_file"],
            "code_no": _extract_code_no(p.get("stm_file", "")),
            "strengths": p.get("strengths", ["N/A"]),
            "test_items": [
                t["name"] for t in p.get("test_items", [])
                if t["name"] != "Description"
                and not (t["name"].startswith("Identification") and not t.get("preparations"))
            ],
        }
        for p in _state["products"]
    ]


@app.get("/api/products/{product_id}")
def get_product(product_id: str):
    product = next((p for p in _state["products"] if p["id"] == product_id), None)
    if not product:
        raise HTTPException(404, "Product not found")
    return product


# ── 계산 ─────────────────────────────────────────────────
class TestItemBatch(BaseModel):
    name: str
    batch_count: int = Field(ge=1, le=100)


class StrengthConfig(BaseModel):
    strength: str
    test_items: list[TestItemBatch]


class CalculateRequest(BaseModel):
    product_id: str
    strength_configs: list[StrengthConfig]


@app.post("/api/calculate")
def calculate(req: CalculateRequest):
    product = next((p for p in _state["products"] if p["id"] == req.product_id), None)
    if not product:
        raise HTTPException(404, "Product not found")
    if not req.strength_configs:
        raise HTTPException(400, "함량 및 시험항목을 선택하세요")

    results = [
        calculate_resources(product, sc.strength, [tb.name], tb.batch_count)
        for sc in req.strength_configs
        for tb in sc.test_items
    ]
    strength_configs_dict = [
        {"strength": sc.strength, "test_items": [{"name": tb.name, "batch_count": tb.batch_count} for tb in sc.test_items]}
        for sc in req.strength_configs
    ]
    merged = merge_all_results(results, strength_configs_dict)
    merged["standards"] = lookup_standards(merged.get("standard_names", []))

    selected_test_names = {tb.name for sc in req.strength_configs for tb in sc.test_items}
    # Uniformity 선택 시 Assay의 column_spec도 포함 (동일 HPLC 조건 사용)
    assay_col_spec = next(
        (item.get("hplc_conditions", {}).get("column_spec")
         for item in product.get("test_items", [])
         if re.match(r"^assay$", item.get("name", ""), re.IGNORECASE)
         and item.get("hplc_conditions", {}).get("column_spec")),
        None,
    )
    col_specs_set: set[str] = set()
    for item in product.get("test_items", []):
        if item.get("name") not in selected_test_names:
            continue
        spec = item.get("hplc_conditions", {}).get("column_spec")
        if spec:
            col_specs_set.add(spec)
        elif re.match(r"uniformity", item.get("name", ""), re.IGNORECASE) and assay_col_spec:
            col_specs_set.add(assay_col_spec)
    merged["columns"] = lookup_columns(list(col_specs_set))
    return merged


# ── 시약 스케일 계산 ──────────────────────────────────────
class ScaleRequest(BaseModel):
    ingredients: list[dict]
    volume_per_batch_ml: float
    actual_prep_ml: float


@app.post("/api/scale")
def scale(req: ScaleRequest):
    return scale_ingredients(req.ingredients, req.volume_per_batch_ml, req.actual_prep_ml)


# ── 파싱 (관리) ──────────────────────────────────────────
@app.get("/api/parse/status")
def parse_status():
    return {
        "parsing": _state["parsing"],
        "error": _state["parse_error"],
        "log": _state["parse_log"][-20:],
        "product_count": len(_state["products"]),
    }


@app.post("/api/parse")
def trigger_parse(background_tasks: BackgroundTasks):
    with _state_lock:
        if _state["parsing"]:
            return {"message": "이미 파싱 중입니다"}
        _state["parsing"] = True
        _state["parse_error"] = None
        _state["parse_log"] = []

    def _log(msg: str):
        print(msg)
        _state["parse_log"].append(msg)

    def _do_parse():
        try:
            products = build_knowledge_base(progress_callback=_log)
            with _state_lock:
                _state["products"] = products
            _log(f"완료: {len(products)}개 제품 파싱됨")
        except Exception as e:
            with _state_lock:
                _state["parse_error"] = str(e)
            _log(f"오류: {e}")
        finally:
            with _state_lock:
                _state["parsing"] = False

    background_tasks.add_task(_do_parse)
    return {"message": "파싱을 시작했습니다"}


# ── API 키 설정 ──────────────────────────────────────────
@app.get("/api/config")
def get_config():
    key = _get_api_key()
    return {"api_key_set": bool(key)}


class ConfigRequest(BaseModel):
    api_key: str


@app.post("/api/config")
def set_config(req: ConfigRequest):
    cfg = _load_config()
    cfg["api_key"] = req.api_key
    _save_config(cfg)
    return {"message": "저장됨"}


# ── 지식베이스 수동 편집 ──────────────────────────────────
@app.put("/api/products/{product_id}")
def update_product(product_id: str, body: dict):
    products = _state["products"]
    idx = next((i for i, p in enumerate(products) if p["id"] == product_id), None)
    if idx is None:
        raise HTTPException(404, "Product not found")
    body["id"] = product_id
    products[idx] = body
    save_knowledge_base(products)
    return {"message": "저장됨"}


# ── 정적 파일 서빙 ────────────────────────────────────────
static_dir = Path(__file__).parent / "static"
app.mount("/", StaticFiles(directory=static_dir, html=True), name="static")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=False)
