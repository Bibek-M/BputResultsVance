"""
BPUT Results Scraper & Excel Exporter
Entry point script.
"""

import sys

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from cli import run_cli

if __name__ == "__main__":
    run_cli()


