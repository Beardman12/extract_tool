from __future__ import annotations

import argparse
import imaplib
import json
import os
import re
import subprocess
import sys
import time
from datetime import datetime
from email import message_from_bytes
from email.header import decode_header
from pathlib import Path
from typing import Dict, Optional, Set, Tuple


POLL_INTERVAL_SECONDS = 300
TARGET_SUBJECT_KEYWORD = "直发报价调整"
DEFAULT_IMAP_HOST = "imap.exmail.qq.com"
DEFAULT_IMAP_PORT = 993

BASE_DIR = Path(__file__).resolve().parents[4]
ENV_PATH = BASE_DIR / ".env"
OUTPUT_ROOT = BASE_DIR / "output" / "mail_attachments" / "direct_price_adjustment"
STATE_PATH = BASE_DIR / "output" / "mail_fetch_state.json"


def load_env_file(env_path: Path) -> Dict[str, str]:
    values: Dict[str, str] = {}
    if not env_path.exists():
        return values

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        values[key] = value
    return values


def get_config() -> Dict[str, str]:
    env_from_file = load_env_file(ENV_PATH)

    def pick(name: str, default: Optional[str] = None) -> str:
        return os.getenv(name) or env_from_file.get(name) or (default or "")

    config = {
        "username": pick("IMAP_USERNAME") or pick("MAIL_USERNAME") or pick("EMAIL_USERNAME"),
        "password": pick("IMAP_PASSWORD") or pick("MAIL_PASSWORD") or pick("EMAIL_PASSWORD"),
        "imap_host": pick("IMAP_HOST", DEFAULT_IMAP_HOST),
        "imap_port": pick("IMAP_PORT", str(DEFAULT_IMAP_PORT)),
        "imap_folder": pick("IMAP_FOLDER", "INBOX"),
        "subject_keyword": pick("TARGET_SUBJECT_KEYWORD", TARGET_SUBJECT_KEYWORD),
    }

    if not config["username"] or not config["password"]:
        raise ValueError(
            "未找到邮箱账号信息，请在 .env 中配置 IMAP_USERNAME/IMAP_PASSWORD "
            "(或 MAIL_USERNAME/MAIL_PASSWORD)。"
        )

    return config


def load_seen_uids() -> Set[str]:
    if not STATE_PATH.exists():
        return set()
    try:
        data = json.loads(STATE_PATH.read_text(encoding="utf-8"))
        return set(data.get("seen_uids", []))
    except Exception:
        return set()


def save_seen_uids(seen_uids: Set[str]) -> None:
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    payload = {"seen_uids": sorted(seen_uids)}
    STATE_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def make_timestamp_folder() -> str:
    now = datetime.now()
    millisecond = int(now.microsecond / 1000)
    return now.strftime("%Y%m%S%H%M") + f"{millisecond:03d}"


def decode_mime_text(value: str) -> str:
    decoded_parts = decode_header(value)
    full_text = ""
    for content, charset in decoded_parts:
        if isinstance(content, bytes):
            full_text += content.decode(charset or "utf-8", errors="replace")
        else:
            full_text += content
    return full_text


def sanitize_filename(filename: str) -> str:
    clean_name = re.sub(r'[\\/:*?"<>|]', "_", filename).strip()
    return clean_name or "attachment.bin"


def decode_filename(part) -> Optional[str]:
    filename = part.get_filename()
    if not filename:
        return None

    try:
        return decode_mime_text(filename)
    except Exception:
        return filename


def save_message_attachments(msg, subject_keyword: str) -> Tuple[int, Optional[Path]]:
    subject = decode_mime_text(str(msg.get("Subject", "")))
    if subject_keyword not in subject:
        return 0, None

    saved_count = 0
    folder_name = make_timestamp_folder()
    target_dir = OUTPUT_ROOT / folder_name
    target_dir.mkdir(parents=True, exist_ok=True)

    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue

        disposition = str(part.get("Content-Disposition", ""))
        if "attachment" not in disposition.lower():
            continue

        filename = decode_filename(part)
        if not filename:
            filename = f"attachment_{saved_count + 1}.bin"
        filename = sanitize_filename(filename)

        payload = part.get_payload(decode=True)
        if payload is None:
            continue

        save_path = target_dir / filename
        # 避免同名覆盖
        if save_path.exists():
            stem = save_path.stem
            suffix = save_path.suffix
            idx = 1
            while True:
                candidate = target_dir / f"{stem}_{idx}{suffix}"
                if not candidate.exists():
                    save_path = candidate
                    break
                idx += 1

        save_path.write_bytes(payload)
        saved_count += 1
        print(f"已保存附件: {save_path}")

    if saved_count == 0 and target_dir.exists() and not any(target_dir.iterdir()):
        target_dir.rmdir()
        target_dir = None

    return saved_count, target_dir


def trigger_prebuild_for_batch(batch_dir: Path) -> None:
    promo_found = any(p.is_file() and p.suffix.lower() == ".xlsx" and "推广报价表" in p.name for p in batch_dir.glob("*.xlsx"))
    vip_found = any(p.is_file() and p.suffix.lower() == ".xlsx" and "直发产品定价" in p.name for p in batch_dir.glob("*.xlsx"))
    if not (promo_found and vip_found):
        print(f"[预生成] 跳过，批次目录内推广/VIP文件不完整: {batch_dir}")
        return

    script_cmd = [
        sys.executable,
        "-m",
        "myproject.scripts.prebuild_batch_outputs",
        "--batch-dir",
        str(batch_dir),
    ]
    child_env = os.environ.copy()
    src_dir = str((BASE_DIR / "src").resolve())
    child_env["PYTHONPATH"] = src_dir if not child_env.get("PYTHONPATH") else src_dir + os.pathsep + child_env["PYTHONPATH"]
    print(f"[预生成] 开始执行: {' '.join(script_cmd)}")
    completed = subprocess.run(script_cmd, cwd=str(BASE_DIR), env=child_env)
    if completed.returncode != 0:
        print(f"[预生成] 执行失败，退出码: {completed.returncode}")
    else:
        print("[预生成] 执行完成")


def fetch_once(config: Dict[str, str], seen_uids: Set[str]) -> int:
    host = config["imap_host"]
    port = int(config["imap_port"])
    username = config["username"]
    password = config["password"]
    folder = config["imap_folder"]
    subject_keyword = config["subject_keyword"]

    print(f"连接 IMAP: {host}:{port}, 文件夹: {folder}")
    with imaplib.IMAP4_SSL(host, port) as client:
        client.login(username, password)
        status, _ = client.select(folder)
        if status != "OK":
            raise RuntimeError(f"无法选择邮箱文件夹: {folder}")

        status, msg_data = client.uid("SEARCH", None, "ALL")
        if status != "OK":
            raise RuntimeError("搜索邮件失败")

        all_uids = msg_data[0].decode("utf-8").split()
        new_uids = [uid for uid in all_uids if uid not in seen_uids]
        print(f"本次扫描 UID 总数: {len(all_uids)}, 未处理: {len(new_uids)}")

        saved_total = 0
        for uid in new_uids:
            status, fetched = client.uid("FETCH", uid, "(RFC822)")
            if status != "OK" or not fetched or fetched[0] is None:
                continue

            raw_email = fetched[0][1]
            if not raw_email:
                continue

            msg = message_from_bytes(raw_email)
            saved, batch_dir = save_message_attachments(msg, subject_keyword)
            saved_total += saved
            seen_uids.add(uid)

            if saved > 0 and batch_dir is not None:
                trigger_prebuild_for_batch(batch_dir)

        return saved_total


def run_single_cycle() -> int:
    config = get_config()
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    seen_uids = load_seen_uids()
    saved_count = fetch_once(config, seen_uids)
    save_seen_uids(seen_uids)
    print(f"本轮完成，保存附件数量: {saved_count}")
    return saved_count


def run_forever(interval_seconds: int = POLL_INTERVAL_SECONDS) -> None:
    print(f"邮件附件定时抓取已启动，每 {interval_seconds} 秒执行一次。")

    while True:
        try:
            run_single_cycle()
        except Exception as exc:
            print(f"抓取发生错误: {exc}")

        time.sleep(interval_seconds)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="IMAP attachment fetch service")
    parser.add_argument("--once", action="store_true", help="Run exactly one fetch cycle")
    parser.add_argument("--interval", type=int, default=POLL_INTERVAL_SECONDS, help="Polling interval in seconds")
    return parser


def main() -> None:
    args = build_parser().parse_args()
    if args.once:
        run_single_cycle()
        return
    run_forever(interval_seconds=max(1, int(args.interval)))


if __name__ == "__main__":
    main()
