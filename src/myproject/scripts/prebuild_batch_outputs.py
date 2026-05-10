#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Dict, List, Optional

from myproject.services.extract_pricing_details import PricingExtractor
from myproject.services.grade_quote_extractor import GradeQuoteExtractor


GRADES: List[str] = ["A", "B", "C", "D", "E", "F"]
PROMO_KEYWORD = "推广报价表"
VIP_KEYWORD = "直发产品定价"


def _is_excel_candidate(p: Path) -> bool:
    return p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")


def _find_latest_xlsx_in_dir(folder: Path, keyword: str) -> Optional[Path]:
    candidates = [p for p in folder.glob("*.xlsx") if _is_excel_candidate(p) and keyword in p.name]
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


def _resolve_files(batch_dir: Path) -> tuple[Path, Path]:
    promo = _find_latest_xlsx_in_dir(batch_dir, PROMO_KEYWORD)
    vip = _find_latest_xlsx_in_dir(batch_dir, VIP_KEYWORD)
    if promo is None:
        raise FileNotFoundError(f"批次目录中未找到推广报价Excel: {batch_dir}")
    if vip is None:
        raise FileNotFoundError(f"批次目录中未找到VIP等级报价Excel: {batch_dir}")
    return promo, vip


def _build_index(promo_file: Path, source_tag: str, rebuild: bool) -> tuple[Path, Dict[str, object]]:
    extractor = PricingExtractor(promo_file)
    index_file = Path("output") / "index" / source_tag / f"{promo_file.stem}_code_sheet_index.json"
    index_file.parent.mkdir(parents=True, exist_ok=True)

    if rebuild or not index_file.exists():
        index_data = extractor.build_code_sheet_index()
        extractor.save_code_sheet_index(index_file, index_data)
    else:
        index_data = extractor.load_code_sheet_index(index_file)
    return index_file, index_data


def _build_structured(vip_file: Path, source_tag: str, rebuild: bool) -> tuple[Path, Dict[str, object]]:
    structured_file = Path("output") / "cache" / source_tag / f"{vip_file.stem}_vip_structured.json"
    structured_file.parent.mkdir(parents=True, exist_ok=True)

    if rebuild or not structured_file.exists():
        extractor = GradeQuoteExtractor(vip_file, sheet_name="等级报价")
        result = extractor.extract()
        structured_file.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    else:
        result = json.loads(structured_file.read_text(encoding="utf-8"))
    return structured_file, result


def _run_query_for_code_grade(
    code: str,
    grade: str,
    promo_file: Path,
    vip_file: Path,
    source_tag: str,
    index_file: Path,
    structured_file: Path,
    snapshot_engine: str,
    output_dir: Path,
    force_regenerate: bool,
) -> int:
    cmd = [
        sys.executable,
        "-m",
        "myproject.scripts.query_product_code",
        "--code",
        code,
        "--grade",
        grade,
        "--promo-input",
        str(promo_file),
        "--vip-input",
        str(vip_file),
        "--promo-index-file",
        str(index_file),
        "--vip-structured-json",
        str(structured_file),
        "--source-tag",
        source_tag,
        "--snapshot-engine",
        snapshot_engine,
        "--output-dir",
        str(output_dir),
    ]
    if force_regenerate:
        cmd.append("--force-regenerate")

    child_env = os.environ.copy()
    src_dir = str((Path.cwd() / "src").resolve())
    child_env["PYTHONPATH"] = src_dir if not child_env.get("PYTHONPATH") else src_dir + os.pathsep + child_env["PYTHONPATH"]
    completed = subprocess.run(cmd, cwd=str(Path.cwd()), env=child_env)
    return completed.returncode


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="按抓取批次预生成所有产品代码及等级输出")
    parser.add_argument("--batch-dir", type=Path, required=True, help="抓取批次目录，例如 output/mail_attachments/direct_price_adjustment/202605101423001")
    parser.add_argument("--output-dir", type=Path, default=Path("output") / "code_query", help="查询输出根目录")
    parser.add_argument("--snapshot-engine", choices=["com", "draw"], default="com", help="截图引擎")
    parser.add_argument("--rebuild-index", action="store_true", help="强制重建 code_sheet_index.json")
    parser.add_argument("--rebuild-structured", action="store_true", help="强制重建 structured.json")
    parser.add_argument("--force-regenerate", action="store_true", help="强制重建 code+grade 结果")
    return parser


def main() -> None:
    args = build_parser().parse_args()
    batch_dir = args.batch_dir.resolve()
    source_tag = batch_dir.name

    promo_file, vip_file = _resolve_files(batch_dir)
    print(f"[预生成] 批次目录: {batch_dir}")
    print(f"[预生成] 推广文件: {promo_file}")
    print(f"[预生成] VIP文件: {vip_file}")

    index_file, index_data = _build_index(promo_file, source_tag, rebuild=args.rebuild_index)
    structured_file, _ = _build_structured(vip_file, source_tag, rebuild=args.rebuild_structured)
    print(f"[预生成] 索引文件: {index_file}")
    print(f"[预生成] 结构化文件: {structured_file}")

    code_to_sheets = index_data.get("code_to_sheets", {}) if isinstance(index_data, dict) else {}
    codes = sorted([str(c).strip().upper() for c in code_to_sheets.keys() if str(c).strip()])
    print(f"[预生成] 产品代码数量: {len(codes)}")

    total = 0
    failed = 0
    for code in codes:
        for grade in GRADES:
            total += 1
            rc = _run_query_for_code_grade(
                code=code,
                grade=grade,
                promo_file=promo_file,
                vip_file=vip_file,
                source_tag=source_tag,
                index_file=index_file,
                structured_file=structured_file,
                snapshot_engine=args.snapshot_engine,
                output_dir=args.output_dir,
                force_regenerate=args.force_regenerate,
            )
            if rc != 0:
                failed += 1
                print(f"[预生成] 失败 code={code} grade={grade} rc={rc}")

    print(f"[预生成] 完成 total={total} failed={failed}")


if __name__ == "__main__":
    main()
