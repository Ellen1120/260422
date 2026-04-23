"""
STM 자원 계산기 - FastAPI 서버
"""
from __future__ import annotations

import json
import os
import threading
from pathlib import Path

from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

from parser import build_knowledge_base, load_knowledge_base, save_knowledge_base
from calculator import calculate_resources, scale_ingredients

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


@app.on_event("startup")
async def startup():
    _state["products"] = load_knowledge_base()
    print(f"Loaded {len(_state['products'])} products from knowledge base.")


# ── 제품 목록 ─────────────────────────────────────────────
@app.get("/api/products")
def get_products():
    return [
        {
            "id": p["id"],
            "name": p["product_name"],
            "stm_file": p["stm_file"],
            "strengths": p.get("strengths", ["N/A"]),
            "test_items": [t["name"] for t in p.get("test_items", [])],
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
class CalculateRequest(BaseModel):
    product_id: str
    strength: str
    test_items: list[str]
    batch_count: int = Field(ge=1, le=100)


@app.post("/api/calculate")
def calculate(req: CalculateRequest):
    product = next((p for p in _state["products"] if p["id"] == req.product_id), None)
    if not product:
        raise HTTPException(404, "Product not found")
    if not req.test_items:
        raise HTTPException(400, "시험항목을 하나 이상 선택하세요")
    return calculate_resources(product, req.strength, req.test_items, req.batch_count)


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
