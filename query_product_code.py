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
import shutil
import traceback
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import load_workbook

from extract_pricing_details import PricingExtractor, ProductLineTable
from grade_quote_extractor import GradeQuoteExtractor, ProductCodeBlock


def _log(msg: str) -> None:
    print(msg, flush=True)


def _find_first_xlsx(workspace: Path, keyword: str) -> Optional[Path]:
    for p in workspace.glob("*.xlsx"):
        if p.name.startswith("~$"):
            continue
        if keyword in p.name:
            return p
    return None


def _normalize_code(code: str) -> str:
    return code.strip().upper()


def _normalize_text(text: Any) -> str:
    if text is None:
        return ""
    s = str(text).strip().upper()
    s = s.replace("（", "(").replace("）", ")")
    s = "".join(s.split())
    return s


def _is_empty_value(v: Any) -> bool:
    return v is None or str(v).strip() == ""


def _is_follow_public(v: Any) -> bool:
    return "随公开价" in str(v)


def _split_country_zone(zone: str) -> List[str]:
    if not zone:
        return []
    tmp = zone.replace("\n", ",").replace("，", ",").replace("/", ",").replace(";", ",")
    tokens = [t.strip().upper() for t in tmp.split(",") if t.strip()]
    return tokens


def _country_aliases(country: str) -> List[str]:
    c = _normalize_text(country)
    aliases = {c}
    mapping = {
        "美国": ["US", "USA", "美国"],
        "德国": ["DE", "GERMANY", "德国"],
        "法国": ["FR", "FRANCE", "法国"],
        "西班牙": ["ES", "SPAIN", "西班牙"],
        "葡萄牙": ["PT", "PORTUGAL", "葡萄牙"],
        "荷兰": ["NL", "NETHERLANDS", "荷兰"],
        "英国": ["GB", "UK", "UNITEDKINGDOM", "英国"],
    }
    for k, vals in mapping.items():
        if k in country:
            aliases.update(vals)
    return list(aliases)


def _country_match(promo_country: str, zone_text: str) -> bool:
    zone_norm = _normalize_text(zone_text)
    if "所有国家" in zone_norm:
        return True

    aliases = set(_country_aliases(promo_country))
    tokens = _split_country_zone(zone_text)
    if not tokens:
        return False

    for token in tokens:
        token_norm = _normalize_text(token)
        if token_norm in aliases:
            return True
        if any(a in token_norm or token_norm in a for a in aliases):
            return True
    return False


def _select_grade_fee(vip_record: Dict[str, Any], grade: str) -> Tuple[Any, Any]:
    key = f"{grade}等级"
    data = vip_record.get(key)
    if not isinstance(data, dict):
        return "", ""
    return data.get("运费"), data.get("处理费")


def _pick_best_vip_record(promo_country: str, promo_weight: str, vip_records: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    w_norm = _normalize_text(promo_weight)
    if not w_norm:
        return None

    exact: List[Dict[str, Any]] = []
    all_country: List[Dict[str, Any]] = []

    for rec in vip_records:
        rec_weight = _normalize_text(rec.get("重量段KG", ""))
        if rec_weight != w_norm:
            continue

        zone = str(rec.get("国家分区", ""))
        if "所有国家" in _normalize_text(zone):
            all_country.append(rec)
            continue

        if _country_match(promo_country, zone):
            exact.append(rec)

    if exact:
        return exact[0]
    if all_country:
        return all_country[0]
    return None


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
    index_file: Optional[Path] = None,
    rebuild_index: bool = False,
) -> Dict[str, Any]:
    _log(f"[推广] 开始读取: {promo_file}")
    _log("[推广] 初始化提取器...")
    extractor = PricingExtractor(promo_file)
    _log("[推广] 提取器初始化完成，开始目录匹配...")
    promo_index_file = index_file or extractor.default_index_file()
    matched_sheets = extractor.get_sheets_by_code(code=code, index_file=promo_index_file, rebuild=rebuild_index)
    if not matched_sheets:
        return {
            "input_excel": str(promo_file),
            "index_file": str(promo_index_file),
            "matched_sheets": [],
            "matched_count": 0,
            "engine": snapshot_engine,
            "snapshots": [],
            "tables": [],
            "message": f"产品代码 {code} 在推广报价中匹配不到工作表",
        }

    _log(f"[推广] 目录命中工作表: {', '.join(matched_sheets)}")
    tables = extractor.extract(include_sheets=matched_sheets)

    matched = []
    for t in tables:
        codes = [c.upper() for c in t.product_codes]
        if code in codes:
            matched.append(t)

    debug_codes = sorted({c.upper() for t in tables for c in t.product_codes})

    snapshot_dir = out_dir / "promo_snapshots"
    snapshots: List[str] = []
    used_engine = snapshot_engine

    if matched:
        _log(f"[推广] 命中产品线数量: {len(matched)}，开始生成截图（{snapshot_engine}）")
        if snapshot_engine == "com":
            try:
                paths = extractor.save_snapshots_com(matched, snapshot_dir)
            except Exception:
                _log("[推广] COM截图失败，自动回退 draw")
                used_engine = "draw"
                paths = extractor.save_snapshots(matched, snapshot_dir)
        else:
            paths = extractor.save_snapshots(matched, snapshot_dir)

        snapshots = [str(p) for p in paths]

    _log(f"[推广] 完成，匹配 {len(matched)}，截图 {len(snapshots)}")

    message = ""
    if not matched:
        if matched_sheets:
            message = (
                f"产品代码 {code} 已命中工作表({', '.join(matched_sheets)})，"
                "但未在可识别的报价表块中匹配到该代码。"
            )
        else:
            message = f"产品代码 {code} 在推广报价中匹配不到工作表"

    return {
        "input_excel": str(promo_file),
        "index_file": str(promo_index_file),
        "matched_sheets": matched_sheets,
        "extracted_table_count": len(tables),
        "extracted_codes": debug_codes,
        "matched_count": len(matched),
        "engine": used_engine,
        "snapshots": snapshots,
        "tables": [_promo_table_to_dict(t) for t in matched],
        "message": message,
    }


def query_vip(
    code: str,
    grade: str,
    vip_file: Path,
    snapshot_engine: str,
    out_dir: Path,
) -> Dict[str, Any]:
    _log(f"[VIP] 开始读取: {vip_file}")
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
    _log(f"[VIP] 命中记录 {len(records)}，命中块 {len(blocks_for_code)}，开始生成截图（{snapshot_engine}）")
    snapshot_items = extractor.save_block_snapshots(
        blocks=blocks_for_code,
        header_row=header_row,
        cm=cm,
        output_dir=snapshot_dir,
        engine=snapshot_engine,
    )

    _log(f"[VIP] 完成，截图 {len(snapshot_items)}")

    return {
        "input_excel": str(vip_file),
        "grade": grade,
        "matched_record_count": len(records),
        "matched_block_count": len(blocks_for_code),
        "records": records,
        "by_product_code": by_code,
        "snapshots": snapshot_items,
    }


def apply_grade_to_promo(
    code: str,
    grade: str,
    promo_file: Path,
    vip_records: List[Dict[str, Any]],
    snapshot_engine: str,
    out_dir: Path,
    include_sheets: Optional[List[str]] = None,
) -> Dict[str, Any]:
    _log(f"[覆盖] 开始按等级 {grade} 覆盖推广报价")
    source_extractor = PricingExtractor(promo_file)
    tables = [
        t
        for t in source_extractor.extract(include_sheets=include_sheets)
        if code in [x.upper() for x in t.product_codes]
    ]

    update_logs: List[Dict[str, Any]] = []
    for table in tables:
        for row in table.rows:
            row_no = int(row.get("行号", 0))
            promo_country = str(row.get("国家", ""))
            promo_weight = str(row.get("重量段/KG", ""))

            rec = _pick_best_vip_record(promo_country, promo_weight, vip_records)
            current_f = row.get("运费（RMB/KG）")
            current_h = row.get("处理费(RMB/票)")

            new_f = ""
            new_h = ""
            status = "unmatched_blank"
            matched_zone = ""

            if rec is not None:
                fee_f, fee_h = _select_grade_fee(rec, grade)
                matched_zone = str(rec.get("国家分区", ""))
                status = "matched"

                if _is_follow_public(fee_f):
                    new_f = current_f
                    status = "follow_public_keep"
                elif _is_empty_value(fee_f):
                    new_f = ""
                else:
                    new_f = fee_f

                if _is_follow_public(fee_h):
                    new_h = current_h
                    status = "follow_public_keep"
                elif _is_empty_value(fee_h):
                    new_h = ""
                else:
                    new_h = fee_h

            update_logs.append(
                {
                    "sheet": table.sheet_name,
                    "产品线": table.product_line,
                    "row": row_no,
                    "国家": promo_country,
                    "重量段KG": promo_weight,
                    "原运费": current_f,
                    "原处理费": current_h,
                    "新运费": new_f,
                    "新处理费": new_h,
                    "匹配状态": status,
                    "匹配国家分区": matched_zone,
                }
            )

    # 在复制文件上覆盖值，保留原文件不变
    modified_excel = out_dir / f"{code}_promo_modified_by_{grade}.xlsx"
    shutil.copy2(promo_file, modified_excel)
    wb = load_workbook(modified_excel)

    for log in update_logs:
        ws = wb[log["sheet"]]
        row_no = int(log["row"])

        target_table = None
        for t in tables:
            if t.sheet_name == log["sheet"] and t.product_line == log["产品线"]:
                target_table = t
                break
        if target_table is None:
            continue

        freight_col = target_table.left_col + 2
        handling_col = target_table.left_col + 3

        ws.cell(row=row_no, column=freight_col).value = log["新运费"]
        ws.cell(row=row_no, column=handling_col).value = log["新处理费"]

    wb.save(modified_excel)
    _log(f"[覆盖] 已写入覆盖结果Excel: {modified_excel}")

    # 用修改后的文件重新提取并截图，保证结构化数据和图片一致
    modified_extractor = PricingExtractor(modified_excel)
    modified_tables_all = modified_extractor.extract(include_sheets=include_sheets)
    modified_tables = [t for t in modified_tables_all if code in [x.upper() for x in t.product_codes]]

    snapshot_dir = out_dir / "promo_modified_snapshots"
    used_engine = snapshot_engine
    _log(f"[覆盖] 开始生成覆盖后截图（{snapshot_engine}）")
    if snapshot_engine == "com":
        try:
            snap_paths = modified_extractor.save_snapshots_com(modified_tables, snapshot_dir)
        except Exception:
            _log("[覆盖] COM截图失败，自动回退 draw")
            used_engine = "draw"
            snap_paths = modified_extractor.save_snapshots(modified_tables, snapshot_dir)
    else:
        snap_paths = modified_extractor.save_snapshots(modified_tables, snapshot_dir)

    _log(f"[覆盖] 完成，截图 {len(snap_paths)}")

    return {
        "grade": grade,
        "modified_excel": str(modified_excel),
        "matched_table_count": len(modified_tables),
        "update_logs": update_logs,
        "modified_tables": [_promo_table_to_dict(t) for t in modified_tables],
        "snapshots": [str(p) for p in snap_paths],
        "engine": used_engine,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="按产品代码查询推广报价和VIP等级报价")
    parser.add_argument("--code", type=str, required=True, help="产品代码，如 CTD")
    parser.add_argument("--grade", choices=["A", "B", "C", "D", "E", "F"], required=True, help="等级报价档位")
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
    parser.add_argument(
        "--promo-index-file",
        type=Path,
        default=None,
        help="推广报价产品代码-工作表目录文件路径；为空时使用默认路径",
    )
    parser.add_argument(
        "--rebuild-promo-index",
        action="store_true",
        help="强制重建推广报价产品代码-工作表目录",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    code = _normalize_code(args.code)
    grade = _normalize_code(args.grade)

    workspace = Path.cwd()

    promo_file = args.promo_input or _find_first_xlsx(workspace, "出口易物流推广报价表")
    vip_file = args.vip_input or _find_first_xlsx(workspace, "2025年直发产品定价+vip")

    if promo_file is None:
        raise FileNotFoundError("未找到推广报价Excel，请使用 --promo-input 指定")
    if vip_file is None:
        raise FileNotFoundError("未找到VIP等级报价Excel，请使用 --vip-input 指定")

    out_dir = args.output_dir / code
    out_dir.mkdir(parents=True, exist_ok=True)

    _log(f"[开始] 产品代码={code}, 等级={grade}, 截图引擎={args.snapshot_engine}")

    promo_result = query_promo(
        code,
        promo_file,
        args.snapshot_engine,
        out_dir,
        index_file=args.promo_index_file,
        rebuild_index=args.rebuild_promo_index,
    )

    if promo_result.get("matched_count", 0) <= 0:
        stop_message = promo_result.get("message") or f"产品代码 {code} 在推广报价中匹配不到工作表，已停止后续执行。"
        payload = {
            "产品代码": code,
            "等级": grade,
            "推广报价": promo_result,
            "message": stop_message,
        }
        output_json = out_dir / f"{code}_query_result.json"
        output_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        _log(payload["message"])
        _log(f"输出JSON: {output_json}")
        return

    vip_result = query_vip(code, grade, vip_file, args.snapshot_engine, out_dir)
    modified_result = apply_grade_to_promo(
        code=code,
        grade=grade,
        promo_file=promo_file,
        vip_records=vip_result["records"],
        snapshot_engine=args.snapshot_engine,
        out_dir=out_dir,
        include_sheets=promo_result.get("matched_sheets"),
    )

    payload = {
        "产品代码": code,
        "等级": grade,
        "推广报价": promo_result,
        "VIP等级报价": vip_result,
        "推广报价_按等级覆盖后": modified_result,
    }

    output_json = out_dir / f"{code}_query_result.json"
    output_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"产品代码: {code}")
    print(f"等级: {grade}")
    print(f"推广索引文件: {promo_result.get('index_file', '')}")
    print(f"推广命中工作表: {', '.join(promo_result.get('matched_sheets', []))}")
    print(f"推广报价匹配: {promo_result['matched_count']}")
    print(f"VIP记录匹配: {vip_result['matched_record_count']}")
    print(f"VIP块匹配: {vip_result['matched_block_count']}")
    print(f"覆盖后推广表匹配: {modified_result['matched_table_count']}")
    print(f"输出JSON: {output_json}")
    print(f"推广截图目录: {out_dir / 'promo_snapshots'}")
    print(f"VIP截图目录: {out_dir / 'vip_snapshots'}")
    print(f"覆盖后推广截图目录: {out_dir / 'promo_modified_snapshots'}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        _log(f"[错误] {type(exc).__name__}: {exc}")
        _log(traceback.format_exc())
        raise
