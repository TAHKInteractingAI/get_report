"""
Entry point for Microsoft Teams Report Dispatcher from Google Sheets.
All core business logic and services are modularized under `src/`.
"""
import gc
from src import (
    email,
    password,
    chat,
    local_tz,
    SPREADSHEET_ID,
    SCOPES,
    MESSAGE_PATTERN,
    CHECKOUT_PATTERN,
    CHECKIN_PATTERN,
    RESET_PATTERN,
    get_hcm_now,
    get_report_time_window,
    _patch_uc_del,
    display_screenshot,
    save_screenshot,
    get_installed_chrome_major_version,
    get_driver,
    get_gspread_client,
    combine_messages,
    preprocess_message,
    is_valid_message,
    _parse_sheet_row_datetime,
    get_filtered_messages,
    write_to_sheet,
    dismiss_teams_popups,
    ensure_chat_tab,
    send_message,
    open_chat,
    login,
    run_report,
)

gc.disable()

if __name__ == "__main__":
    run_report()
