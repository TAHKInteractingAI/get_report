"""
Entry point for Microsoft Teams Report Dispatcher from Google Sheets.
All core business logic and services are modularized under `src/`.
"""
import gc
from src import run_report

gc.disable()

if __name__ == "__main__":
    run_report()
