from __future__ import annotations

import argparse
import time

from myproject.cron.jobs.mail_attachment_fetch import run_single_cycle


DEFAULT_INTERVAL_SECONDS = 300


def main() -> None:
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
