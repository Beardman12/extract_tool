#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import json
import logging
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Literal, Optional

from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field

from extract_pricing_details import PricingExtractor
from grade_quote_extractor import GradeQuoteExtractor


WORKSPACE = Path(__file__).resolve().parent
OUTPUT_DIR = WORKSPACE / "output"
CACHE_DIR = OUTPUT_DIR / "cache"
LOG_DIR = WORKSPACE / "logs"


_logger: Optional[logging.Logger] = None
_logger_date: Optional[str] = None


def get_logger() -> logging.Logger:
    global _logger, _logger_date

    current_date = datetime.now().strftime("%Y%m%d")
    if _logger is not None and _logger_date == current_date:
        return _logger

    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"{current_date}.log"

    logger = logging.getLogger("ai_price_api")
    logger.setLevel(logging.INFO)
    logger.propagate = False

    logger.handlers.clear()
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    _logger = logger
    _logger_date = current_date
    return logger


def _find_first_xlsx(workspace: Path, keyword: str) -> Optional[Path]:
    for p in workspace.glob("*.xlsx"):
        if p.name.startswith("~$"):
            continue
        if keyword in p.name:
            return p
    return None


def _resolve_promo_file(promo_input: Optional[str]) -> Path:
    if promo_input:
        p = Path(promo_input)
        if p.is_absolute():
            return p
        return (WORKSPACE / p).resolve()

    found = _find_first_xlsx(WORKSPACE, "出口易物流推广报价表")
    if found is None:
        raise HTTPException(status_code=404, detail="未找到推广报价Excel，请通过 promo_input 指定")
    return found


def _resolve_vip_file(vip_input: Optional[str]) -> Path:
    if vip_input:
        p = Path(vip_input)
        if p.is_absolute():
            return p
        return (WORKSPACE / p).resolve()

    found = _find_first_xlsx(WORKSPACE, "2025年直发产品定价+vip")
    if found is None:
        raise HTTPException(status_code=404, detail="未找到VIP等级报价Excel，请通过 vip_input 指定")
    return found


def _default_vip_structured_json(vip_file: Path) -> Path:
    return CACHE_DIR / f"{vip_file.stem}_vip_structured.json"


def _ensure_vip_structured_json(
    vip_file: Path,
    structured_json: Optional[str],
    rebuild: bool,
) -> tuple[Path, Dict[str, Any], bool]:
    logger = get_logger()
    structured_file = Path(structured_json) if structured_json else _default_vip_structured_json(vip_file)
    if not structured_file.is_absolute():
        structured_file = (WORKSPACE / structured_file).resolve()

    if rebuild or not structured_file.exists():
        logger.info("VIP结构化JSON不存在或要求重建，开始解析: %s", vip_file)
        extractor = GradeQuoteExtractor(vip_file, sheet_name="等级报价")
        result = extractor.extract()
        structured_file.parent.mkdir(parents=True, exist_ok=True)
        structured_file.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        logger.info("VIP结构化JSON已生成: %s", structured_file)
        return structured_file, result, True

    logger.info("使用已有VIP结构化JSON: %s", structured_file)
    result = json.loads(structured_file.read_text(encoding="utf-8"))
    return structured_file, result, False


class RunQueryRequest(BaseModel):
    code: str = Field(..., description="产品代码，例如 SUX")
    grade: Literal["A", "B", "C", "D", "E", "F"]
    snapshot_engine: Literal["com", "draw"] = "com"
    promo_input: Optional[str] = None
    vip_input: Optional[str] = None
    promo_index_file: Optional[str] = None
    rebuild_promo_index: bool = False
    country_match_config: Optional[str] = None
    vip_structured_json: Optional[str] = None
    rebuild_vip_structured_json: bool = False


app = FastAPI(title="AI Price API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.middleware("http")
async def request_log_middleware(request, call_next):
    logger = get_logger()
    start = time.time()
    logger.info("REQ %s %s", request.method, request.url.path)
    response = await call_next(request)
    cost_ms = int((time.time() - start) * 1000)
    logger.info("RESP %s %s status=%s cost_ms=%s", request.method, request.url.path, response.status_code, cost_ms)
    return response


@app.get("/api/promo/options")
def get_promo_options(
    promo_input: Optional[str] = Query(default=None),
    rebuild_index: bool = Query(default=False),
    include_content: bool = Query(default=True),
):
    logger = get_logger()
    promo_file = _resolve_promo_file(promo_input)
    if not promo_file.exists():
        raise HTTPException(status_code=404, detail=f"推广报价文件不存在: {promo_file}")

    logger.info("读取推广选项，文件=%s rebuild_index=%s include_content=%s", promo_file, rebuild_index, include_content)

    extractor = PricingExtractor(promo_file)
    index_file = extractor.default_index_file()

    if rebuild_index or not index_file.exists():
        index_data = extractor.build_code_sheet_index()
        extractor.save_code_sheet_index(index_file, index_data)
    else:
        index_data = extractor.load_code_sheet_index(index_file)

    code_to_sheets: Dict[str, List[str]] = index_data.get("code_to_sheets", {})

    content_map: Dict[str, List[Dict[str, Any]]] = {}
    if include_content:
        tables = extractor.extract()
        for t in tables:
            item = {
                "sheet": t.sheet_name,
                "product_line": t.product_line,
                "update_time": t.update_time,
                "header_row": t.header_row,
                "row_count": len(t.rows),
                "headers": t.headers,
            }
            for c in t.product_codes:
                key = str(c).strip().upper()
                if not key:
                    continue
                content_map.setdefault(key, []).append(item)

    options: List[Dict[str, Any]] = []
    for code in sorted(code_to_sheets.keys()):
        options.append(
            {
                "code": code,
                "sheets": code_to_sheets.get(code, []),
                "content": content_map.get(code, []),
            }
        )

    return {
        "promo_file": str(promo_file),
        "index_file": str(index_file),
        "code_count": len(options),
        "options": options,
    }


@app.get("/api/vip/options")
def get_vip_options(
    vip_input: Optional[str] = Query(default=None),
    structured_json: Optional[str] = Query(default=None),
    rebuild_structured_json: bool = Query(default=False),
    include_full_data: bool = Query(default=True),
):
    vip_file = _resolve_vip_file(vip_input)
    if not vip_file.exists():
        raise HTTPException(status_code=404, detail=f"VIP报价文件不存在: {vip_file}")

    structured_file, result, rebuilt = _ensure_vip_structured_json(
        vip_file=vip_file,
        structured_json=structured_json,
        rebuild=rebuild_structured_json,
    )

    records = result.get("records", [])
    code_stats: Dict[str, Dict[str, Any]] = {}
    for rec in records:
        code = str(rec.get("产品代码", "")).strip().upper()
        if not code:
            continue
        stat = code_stats.setdefault(
            code,
            {
                "code": code,
                "record_count": 0,
                "zones": set(),
            },
        )
        stat["record_count"] += 1
        zone = str(rec.get("国家分区", "")).strip()
        if zone:
            stat["zones"].add(zone)

    code_options: List[Dict[str, Any]] = []
    for code in sorted(code_stats.keys()):
        item = code_stats[code]
        code_options.append(
            {
                "code": item["code"],
                "record_count": item["record_count"],
                "zones": sorted(list(item["zones"])),
            }
        )

    return {
        "vip_file": str(vip_file),
        "structured_json": str(structured_file),
        "rebuilt": rebuilt,
        "grades": ["A", "B", "C", "D", "E", "F"],
        "code_options": code_options,
        "meta": result.get("meta", {}),
        "structured_data": result if include_full_data else None,
    }


@app.post("/api/query/run")
def run_query(req: RunQueryRequest):
    logger = get_logger()
    code = req.code.strip().upper()
    grade = req.grade.strip().upper()

    script = WORKSPACE / "query_product_code.py"
    if not script.exists():
        raise HTTPException(status_code=404, detail=f"未找到脚本: {script}")

    cmd: List[str] = [
        sys.executable,
        str(script),
        "--code",
        code,
        "--grade",
        grade,
        "--snapshot-engine",
        req.snapshot_engine,
    ]

    if req.promo_input:
        cmd.extend(["--promo-input", req.promo_input])
    if req.vip_input:
        cmd.extend(["--vip-input", req.vip_input])
    if req.promo_index_file:
        cmd.extend(["--promo-index-file", req.promo_index_file])
    if req.rebuild_promo_index:
        cmd.append("--rebuild-promo-index")
    if req.country_match_config:
        cmd.extend(["--country-match-config", req.country_match_config])
    if req.vip_structured_json:
        cmd.extend(["--vip-structured-json", req.vip_structured_json])
    if req.rebuild_vip_structured_json:
        cmd.append("--rebuild-vip-structured-json")

    logger.info("执行查询命令: %s", " ".join(cmd))
    completed = subprocess.run(
        cmd,
        cwd=str(WORKSPACE),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    output_json = OUTPUT_DIR / "code_query" / code / f"{code}_query_result.json"
    payload: Optional[Dict[str, Any]] = None
    if output_json.exists():
        payload = json.loads(output_json.read_text(encoding="utf-8"))

    return {
        "command": cmd,
        "returncode": completed.returncode,
        "stdout": completed.stdout,
        "stderr": completed.stderr,
        "output_json": str(output_json),
        "result": payload,
    }


@app.get("/health")
def health_check():
    return {"status": "ok", "time": datetime.now().isoformat(timespec="seconds")}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("api_server:app", host="0.0.0.0", port=8000, reload=False)
