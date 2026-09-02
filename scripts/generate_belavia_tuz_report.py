#!/usr/bin/env python3
"""Generate Belavia TUZ performance report (same logic as Utair)."""

from generate_utair_tuz_report import main

if __name__ == "__main__":
    main(["--client", "belavia"])
