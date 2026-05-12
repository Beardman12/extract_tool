from __future__ import annotations

import argparse
import atexit
import os
import time
from datetime import datetime
from pathlib import Path

from myproject.cron.jobs.mail_attachment_fetch import run_single_cycle


DEFAULT_INTERVAL_SECONDS = 300
BASE_DIR = Path(__file__).resolve().parents[3]
LOCK_PATH = BASE_DIR / "output" / "scheduler_mail_fetch.lock"
_LOCK_FILE = None


def _acquire_single_instance_lock() -> None:
    global _LOCK_FILE
    LOCK_PATH.parent.mkdir(parents=True, exist_ok=True)

    lock_file = open(LOCK_PATH, "a+", encoding="utf-8")
    try:
        if os.name == "nt":
            import msvcrt

            lock_file.seek(0)
            msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl

            fcntl.flock(lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        lock_file.close()
        raise RuntimeError(f"scheduler 已在运行，锁文件: {LOCK_PATH}")

    lock_file.seek(0)
    lock_file.truncate()
    lock_file.write(f"pid={os.getpid()}\nstarted_at={datetime.now().isoformat()}\n")
    lock_file.flush()
    _LOCK_FILE = lock_file


def _release_single_instance_lock() -> None:
    global _LOCK_FILE
    if _LOCK_FILE is None:
        return

    try:
        if os.name == "nt":
            import msvcrt

            _LOCK_FILE.seek(0)
            msvcrt.locking(_LOCK_FILE.fileno(), msvcrt.LK_UNLCK, 1)
        else:
            import fcntl

            fcntl.flock(_LOCK_FILE.fileno(), fcntl.LOCK_UN)
    except OSError:
        pass
    finally:
        _LOCK_FILE.close()
        _LOCK_FILE = None


def main() -> None:
    _acquire_single_instance_lock()
    atexit.register(_release_single_instance_lock)

    parser = argparse.ArgumentParser(description="Run scheduled jobs")
    parser.add_argument("--once", action="store_true", help="Run all jobs once and exit")
    parser.add_argument("--interval", type=int, default=DEFAULT_INTERVAL_SECONDS, help="Loop interval seconds")
    args = parser.parse_args()

    if args.once:
        try:
            result = run_single_cycle()
            print(f"[scheduler] job=mail_attachment_fetch result={result}")
        except Exception as exc:
            print(f"[scheduler] job=mail_attachment_fetch failed: {type(exc).__name__}: {exc}")
    else:
        interval_seconds = max(1, int(args.interval))
        print(f"[scheduler] started, interval={interval_seconds}s")
        while True:
            try:
                result = run_single_cycle()
                print(f"[scheduler] job=mail_attachment_fetch result={result}")
            except Exception as exc:
                print(f"[scheduler] job=mail_attachment_fetch failed: {type(exc).__name__}: {exc}")
            time.sleep(interval_seconds)


if __name__ == "__main__":
    main()
