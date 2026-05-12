from __future__ import annotations

import argparse
import imaplib
import json
import os
import re
import subprocess
import sys
import time
from datetime import datetime, timedelta, timezone
from email import message_from_bytes
from email.header import decode_header
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple


POLL_INTERVAL_SECONDS = 300
TARGET_SUBJECT_KEYWORD = "直发报价"
DEFAULT_IMAP_HOST = "imap.exmail.qq.com"
DEFAULT_IMAP_PORT = 993
DEFAULT_LOOKBACK_DAYS = 5

def _resolve_base_dir() -> Path:
    # 允许外部显式指定，便于服务部署时固定项目根目录。
    env_root = os.getenv("PROJECT_ROOT", "").strip()
    if env_root:
        candidate = Path(env_root).expanduser().resolve()
        if candidate.exists():
            return candidate

    # 优先从当前文件向上找项目标记。
    file_path = Path(__file__).resolve()
    for parent in file_path.parents:
        if (parent / "pyproject.toml").exists() and (parent / "src").exists():
            return parent

    # 兼容从项目目录启动但模块来源不在项目内（如 site-packages）。
    cwd = Path.cwd().resolve()
    for parent in [cwd, *cwd.parents]:
        if (parent / "pyproject.toml").exists() and (parent / "src").exists():
            return parent

    # 回退到原始推断逻辑，保持向后兼容。
    return file_path.parents[4]


BASE_DIR = _resolve_base_dir()
ENV_PATH = BASE_DIR / ".env"
OUTPUT_ROOT = BASE_DIR / "output" / "mail_attachments" / "direct_price_adjustment"
STATE_PATH = BASE_DIR / "output" / "mail_fetch_state.json"
DEFAULT_STATE_PAYLOAD = {
    "seen_uids": [],
    "seen_message_ids": [],
    "seen_subjects": {},
}


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
        "lookback_days": pick("MAIL_LOOKBACK_DAYS", str(DEFAULT_LOOKBACK_DAYS)),
    }

    if not config["username"] or not config["password"]:
        raise ValueError(
            "未找到邮箱账号信息，请在 .env 中配置 IMAP_USERNAME/IMAP_PASSWORD "
            "(或 MAIL_USERNAME/MAIL_PASSWORD)。"
        )

    return config


def ensure_state_file() -> None:
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    if not STATE_PATH.exists():
        STATE_PATH.write_text(json.dumps(DEFAULT_STATE_PAYLOAD, ensure_ascii=False, indent=2), encoding="utf-8")


def load_seen_state() -> Tuple[Set[str], Set[str], Dict[str, str]]:
    ensure_state_file()
    try:
        data = json.loads(STATE_PATH.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("state payload is not an object")
        seen_uids_raw = data.get("seen_uids", [])
        seen_message_ids_raw = data.get("seen_message_ids", [])
        seen_subjects_raw = data.get("seen_subjects", {})
        if not isinstance(seen_uids_raw, list) or not isinstance(seen_message_ids_raw, list):
            raise ValueError("state payload fields are invalid")
        if not isinstance(seen_subjects_raw, dict):
            seen_subjects_raw = {}
        seen_uids = {str(item) for item in seen_uids_raw if str(item).strip()}
        seen_message_ids = {str(item) for item in seen_message_ids_raw if str(item).strip()}
        seen_subjects: Dict[str, str] = {
            str(uid): str(subject)
            for uid, subject in seen_subjects_raw.items()
            if str(uid).strip()
        }
        return seen_uids, seen_message_ids, seen_subjects
    except Exception:
        # 状态文件损坏时自动重建，避免任务中断。
        STATE_PATH.write_text(json.dumps(DEFAULT_STATE_PAYLOAD, ensure_ascii=False, indent=2), encoding="utf-8")
        return set(), set(), {}


def save_seen_state(seen_uids: Set[str], seen_message_ids: Set[str], seen_subjects: Dict[str, str]) -> None:
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)

    # 与磁盘现有状态做并集，降低并发运行时的覆盖风险。
    disk_uids, disk_message_ids, disk_subjects = load_seen_state()
    merged_uids = set(seen_uids) | disk_uids
    merged_message_ids = set(seen_message_ids) | disk_message_ids
    merged_subjects = dict(disk_subjects)
    merged_subjects.update({k: v for k, v in seen_subjects.items() if k})

    payload = {
        "seen_uids": sorted(merged_uids),
        "seen_message_ids": sorted(merged_message_ids),
        "seen_subjects": {uid: merged_subjects.get(uid, "") for uid in sorted(merged_uids)},
    }
    temp_path = STATE_PATH.with_suffix(STATE_PATH.suffix + ".tmp")
    temp_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    temp_path.replace(STATE_PATH)


def _remember_seen(
    uid: str,
    message_id: str,
    subject: str,
    seen_uids: Set[str],
    seen_message_ids: Set[str],
    seen_subjects: Dict[str, str],
) -> None:
    if uid:
        seen_uids.add(uid)
        if subject and uid not in seen_subjects:
            seen_subjects[uid] = subject
    if message_id:
        seen_message_ids.add(message_id)


def _extract_message_id(msg) -> str:
    return str(msg.get("Message-ID", "")).strip().lower()


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


def _normalize_subject(subject: str) -> str:
    normalized = subject.strip()
    # 去掉常见转发/回复前缀，如 Re: / Fw: / Fwd:
    while True:
        updated = re.sub(r"(?i)^\s*(re|fw|fwd)\s*:\s*", "", normalized).strip()
        if updated == normalized:
            break
        normalized = updated
    return normalized


def _is_target_subject(subject: str, subject_keyword: str) -> bool:
    return subject_keyword in _normalize_subject(subject)


def _collect_required_attachments(msg) -> List[Tuple[str, bytes]]:
    promo_candidate: Optional[Tuple[str, bytes]] = None
    vip_candidate: Optional[Tuple[str, bytes]] = None

    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue

        disposition = str(part.get("Content-Disposition", ""))
        if "attachment" not in disposition.lower():
            continue

        filename = decode_filename(part) or "attachment.bin"
        filename = sanitize_filename(filename)
        payload = part.get_payload(decode=True)
        if payload is None:
            continue

        lower_name = filename.lower()
        if "推广报价" in filename:
            promo_candidate = (filename, payload)
        if "vip" in lower_name:
            vip_candidate = (filename, payload)

    if promo_candidate is None or vip_candidate is None:
        return []
    return [promo_candidate, vip_candidate]


def save_message_attachments(msg) -> Tuple[int, Optional[Path]]:
    required_attachments = _collect_required_attachments(msg)
    if not required_attachments:
        return 0, None

    saved_count = 0
    folder_name = make_timestamp_folder()
    target_dir = OUTPUT_ROOT / folder_name
    target_dir.mkdir(parents=True, exist_ok=True)

    for filename, payload in required_attachments:
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
    promo_found = any(p.is_file() and p.suffix.lower() == ".xlsx" and "推广报价" in p.name for p in batch_dir.glob("*.xlsx"))
    vip_found = any(p.is_file() and p.suffix.lower() == ".xlsx" and "vip" in p.name.lower() for p in batch_dir.glob("*.xlsx"))
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


def _parse_mail_datetime(msg) -> datetime:
    raw_date = str(msg.get("Date", "")).strip()
    if not raw_date:
        return datetime.min.replace(tzinfo=timezone.utc)

    try:
        parsed = parsedate_to_datetime(raw_date)
    except Exception:
        return datetime.min.replace(tzinfo=timezone.utc)

    if parsed.tzinfo is None:
        return parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(timezone.utc)


def _imap_since_date(days: int) -> str:
    # IMAP SINCE 使用格式: 11-May-2026
    target = datetime.now() - timedelta(days=max(0, days))
    return target.strftime("%d-%b-%Y")


def fetch_once(config: Dict[str, str], seen_uids: Set[str], seen_message_ids: Set[str], seen_subjects: Dict[str, str]) -> int:
    host = config["imap_host"]
    port = int(config["imap_port"])
    username = config["username"]
    password = config["password"]
    folder = config["imap_folder"]
    subject_keyword = config["subject_keyword"]
    lookback_days = max(1, int(config.get("lookback_days", DEFAULT_LOOKBACK_DAYS)))
    since_date = _imap_since_date(lookback_days)

    print(f"连接 IMAP: {host}:{port}, 文件夹: {folder}")
    with imaplib.IMAP4_SSL(host, port) as client:
        client.login(username, password)
        status, _ = client.select(folder)
        if status != "OK":
            raise RuntimeError(f"无法选择邮箱文件夹: {folder}")

        status, msg_data = client.uid("SEARCH", None, "SINCE", since_date)
        if status != "OK":
            raise RuntimeError("搜索邮件失败")

        all_uids = msg_data[0].decode("utf-8").split() if msg_data and msg_data[0] else []
        new_uids = [uid for uid in all_uids if uid not in seen_uids]
        print(f"本次扫描 UID 总数(最近{lookback_days}天): {len(all_uids)}, 未处理: {len(new_uids)}")

        if not new_uids:
            return 0

        matching_messages: List[Tuple[datetime, int, str, str, str]] = []

        for uid in new_uids:
            status, fetched = client.uid("FETCH", uid, "(BODY.PEEK[HEADER.FIELDS (SUBJECT DATE MESSAGE-ID)])")
            if status != "OK" or not fetched or fetched[0] is None:
                continue

            raw_header = fetched[0][1]
            if not raw_header:
                continue

            header_msg = message_from_bytes(raw_header)
            message_id = _extract_message_id(header_msg)
            subject = decode_mime_text(str(header_msg.get("Subject", "")))

            if message_id and message_id in seen_message_ids:
                # Message-ID 已处理过时，回填 UID，防止后续轮次重复比较。
                _remember_seen(uid, message_id, subject, seen_uids, seen_message_ids, seen_subjects)
                continue

            if not _is_target_subject(subject, subject_keyword):
                continue

            msg_time = _parse_mail_datetime(header_msg)
            uid_num = int(uid) if uid.isdigit() else 0
            matching_messages.append((msg_time, uid_num, uid, message_id, subject))

        if not matching_messages:
            return 0

        matching_messages.sort(key=lambda x: (x[0], x[1]), reverse=True)
        print(f"匹配主题邮件: {len(matching_messages)}，本轮按时间倒序仅处理第一封附件完整邮件")

        for _, _, uid, header_message_id, subject in matching_messages:
            status, fetched = client.uid("FETCH", uid, "(RFC822)")
            if status != "OK" or not fetched or fetched[0] is None:
                continue

            raw_email = fetched[0][1]
            if not raw_email:
                continue

            msg = message_from_bytes(raw_email)
            message_id = _extract_message_id(msg) or header_message_id
            saved, batch_dir = save_message_attachments(msg)

            # 同一轮只处理一封，且该封必须同时含推广报价与VIP两个附件。
            if saved > 0 and batch_dir is not None:
                # 标记本轮所有命中主题的邮件为已处理，避免后续轮次再回头处理旧邮件。
                for _, _, mark_uid, mark_message_id, mark_subject in matching_messages:
                    _remember_seen(mark_uid, mark_message_id, mark_subject, seen_uids, seen_message_ids, seen_subjects)

                _remember_seen(uid, message_id, subject, seen_uids, seen_message_ids, seen_subjects)
                trigger_prebuild_for_batch(batch_dir)
                return saved

            # 避免同一封不完整邮件在后续轮次反复命中。
            _remember_seen(uid, message_id, subject, seen_uids, seen_message_ids, seen_subjects)

        print("命中主题邮件存在，但未找到同时包含推广报价与VIP附件的邮件")
        return 0


def run_single_cycle() -> int:
    ensure_state_file()
    seen_uids, seen_message_ids, seen_subjects = load_seen_state()
    saved_count = 0
    print(f"状态文件路径: {STATE_PATH}")
    try:
        config = get_config()
        OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
        saved_count = fetch_once(config, seen_uids, seen_message_ids, seen_subjects)
    finally:
        # 即使抓取流程中途异常，也要落盘已标记状态，避免下轮重复处理。
        try:
            save_seen_state(seen_uids, seen_message_ids, seen_subjects)
        except Exception as exc:
            print(f"状态文件写入失败: {STATE_PATH}, error={type(exc).__name__}: {exc}")
            raise
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
