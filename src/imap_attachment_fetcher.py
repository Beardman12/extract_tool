#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

from myproject.cron.jobs.mail_attachment_fetch import main as core_main


def main() -> None:
    core_main()


if __name__ == "__main__":
    main()
