#!/usr/bin/env python3
"""Generate Belavia TUZ performance report (same logic as Utair)."""

import sys

from generate_utair_tuz_report import main

if __name__ == "__main__":
    main(["--client", "belavia", *sys.argv[1:]])
