#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
按产品代码查询两类报价数据（推广报价 + VIP等级报价），并输出结构化数据和截图。

示例:
python query_product_code.py --code CTD
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict, List, Optional

from extract_pricing_details import PricingExtractor, ProductLineTable
from grade_quote_extractor import GradeQuoteExtractor, ProductCodeBlock


def _find_first_xlsx(workspace: Path, keyword: str) -> Optional[Path]:
    for p in workspace.glob("*.xlsx"):
        if p.name.startswith("~$"):
            continue
        if keyword in p.name:
            return p
    return None


def _normalize_code(code: str) -> str:
    return code.strip().upper()


def _promo_table_to_dict(table: ProductLineTable) -> Dict[str, Any]:
    return {
        "sheet": table.sheet_name,
        "产品线": table.product_line,
        "产品代码": table.product_codes,
        "更新时间": table.update_time,
        "header_row": table.header_row,
        "data_start_row": table.data_start_row,
        "data_end_row": table.data_end_row,
        "headers": table.headers,
        "range": {
            "top_left": f"{table.left_col},{table.top_row}",
            "bottom_right": f"{table.right_col},{table.bottom_row}",
        },
        "note": {
            "has_note": table.note_row > 0,
            "row": table.note_row,
            "text": table.note_text,
        },
        "rows": table.rows,
    }


def query_promo(
    code: str,
    promo_file: Path,
    snapshot_engine: str,
    out_dir: Path,
) -> Dict[str, Any]:
    extractor = PricingExtractor(promo_file)
    tables = extractor.extract()

    matched = []
    for t in tables:
        codes = [c.upper() for c in t.product_codes]
        if code in codes:
            matched.append(t)

    snapshot_dir = out_dir / "promo_snapshots"
    snapshots: List[str] = []
    used_engine = snapshot_engine

    if matched:
        if snapshot_engine == "com":
            try:
                paths = extractor.save_snapshots_com(matched, snapshot_dir)
            except Exception:
                used_engine = "draw"
                paths = extractor.save_snapshots(matched, snapshot_dir)
        else:
            paths = extractor.save_snapshots(matched, snapshot_dir)

        snapshots = [str(p) for p in paths]

    return {
        "input_excel": str(promo_file),
        "matched_count": len(matched),
        "engine": used_engine,
        "snapshots": snapshots,
        "tables": [_promo_table_to_dict(t) for t in matched],
    }


def query_vip(
    code: str,
    vip_file: Path,
    snapshot_engine: str,
    out_dir: Path,
) -> Dict[str, Any]:
    extractor = GradeQuoteExtractor(vip_file, sheet_name="等级报价")
    result = extractor.extract()

    records = [r for r in result["records"] if str(r.get("产品代码", "")).upper() == code]
    by_code = result["by_product_code"].get(code)

    blocks_for_code: List[ProductCodeBlock] = []
    for b in result.get("blocks", []):
        code_list = [str(x).upper() for x in b.get("产品代码列表", [])]
        if code in code_list:
            blocks_for_code.append(
                ProductCodeBlock(
                    product_name=b.get("产品名称", ""),
                    service_text=b.get("服务代码原文", ""),
                    service_codes=[code],
                    start_row=int(b.get("start_row", 0)),
                    end_row=int(b.get("end_row", 0)),
                )
            )

    header_row = int(result["meta"]["header_row"])
    cm = extractor._build_column_map(header_row)

    snapshot_dir = out_dir / "vip_snapshots"
    snapshot_items = extractor.save_block_snapshots(
        blocks=blocks_for_code,
        header_row=header_row,
        cm=cm,
        output_dir=snapshot_dir,
        engine=snapshot_engine,
    )

    return {
        "input_excel": str(vip_file),
        "matched_record_count": len(records),
        "matched_block_count": len(blocks_for_code),
        "records": records,
        "by_product_code": by_code,
        "snapshots": snapshot_items,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="按产品代码查询推广报价和VIP等级报价")
    parser.add_argument("--code", type=str, required=True, help="产品代码，如 CTD")
    parser.add_argument(
        "--promo-input",
        type=Path,
        default=None,
        help="推广报价Excel路径（默认自动识别：出口易物流推广报价表）",
    )
    parser.add_argument(
        "--vip-input",
        type=Path,
        default=None,
        help="VIP等级报价Excel路径（默认自动识别：2025年直发产品定价+vip）",
    )
    parser.add_argument(
        "--snapshot-engine",
        choices=["com", "draw"],
        default="com",
        help="截图引擎，默认 com；失败时自动回退 draw",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("output") / "code_query",
        help="输出目录",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    code = _normalize_code(args.code)

    workspace = Path.cwd()

    promo_file = args.promo_input or _find_first_xlsx(workspace, "出口易物流推广报价表")
    vip_file = args.vip_input or _find_first_xlsx(workspace, "2025年直发产品定价+vip")

    if promo_file is None:
        raise FileNotFoundError("未找到推广报价Excel，请使用 --promo-input 指定")
    if vip_file is None:
        raise FileNotFoundError("未找到VIP等级报价Excel，请使用 --vip-input 指定")

    out_dir = args.output_dir / code
    out_dir.mkdir(parents=True, exist_ok=True)

    promo_result = query_promo(code, promo_file, args.snapshot_engine, out_dir)
    vip_result = query_vip(code, vip_file, args.snapshot_engine, out_dir)

    payload = {
        "产品代码": code,
        "推广报价": promo_result,
        "VIP等级报价": vip_result,
    }

    output_json = out_dir / f"{code}_query_result.json"
    output_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"产品代码: {code}")
    print(f"推广报价匹配: {promo_result['matched_count']}")
    print(f"VIP记录匹配: {vip_result['matched_record_count']}")
    print(f"VIP块匹配: {vip_result['matched_block_count']}")
    print(f"输出JSON: {output_json}")
    print(f"推广截图目录: {out_dir / 'promo_snapshots'}")
    print(f"VIP截图目录: {out_dir / 'vip_snapshots'}")


if __name__ == "__main__":
    main()
