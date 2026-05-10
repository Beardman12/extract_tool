#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import sys
from pathlib import Path

SRC_DIR = Path(__file__).resolve().parent / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from myproject.scripts.prebuild_batch_outputs import main as core_main


def main() -> None:
    core_main()


if __name__ == "__main__":
    main()
