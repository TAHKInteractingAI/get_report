import os
import sys
import time
import datetime
from pathlib import Path

if __package__ is None or __package__ == "":
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from src.config import get_hcm_now, get_report_time_window, SPREADSHEET_ID, chat
    from src.browser import save_screenshot
    from src.sheets import (
        get_gspread_client,
        get_filtered_messages,
        combine_messages,
        write_to_sheet,
    )
    from src.teams import login, open_chat, send_message
else:
    from .config import get_hcm_now, get_report_time_window, SPREADSHEET_ID, chat
    from .browser import save_screenshot
    from .sheets import (
        get_gspread_client,
        get_filtered_messages,
        combine_messages,
        write_to_sheet,
    )
    from .teams import login, open_chat, send_message


def run_report():
    utc_now = datetime.datetime.now(datetime.timezone.utc)
    hcm_now = get_hcm_now()
    start_dt, end_dt, shift_name = get_report_time_window(hcm_now)
    ci_trigger_env = os.getenv("CI_TRIGGER_TIME", "N/A (Chạy trực tiếp)")
    ci_event_name = os.getenv("GITHUB_EVENT_NAME", "local_run")

    print("=" * 65)
    print("⏰ KIỂM TRA THỜI GIAN & CA BÁO CÁO (MÚI GIỜ HỒ CHÍ MINH - UTC+7)")
    print(f" - Sự kiện kích hoạt (CI Event) : {ci_event_name}")
    print(f" - Giờ CI nhận lệnh (HCM)       : {ci_trigger_env}")
    print(f" - Giờ máy chủ hiện tại (UTC)   : {utc_now.strftime('%Y-%m-%d %H:%M:%S UTC')}")
    print(f" - Giờ Việt Nam hiện tại (HCM)  : {hcm_now.strftime('%Y-%m-%d %H:%M:%S %Z%z')}")
    print(f" - Ca báo cáo được nhận diện    : {shift_name}")
    print(
        f" - Khung giờ lọc báo cáo        : "
        f"{start_dt.strftime('%Y-%m-%d %H:%M:%S')} -> {end_dt.strftime('%Y-%m-%d %H:%M:%S')} (Asia/Ho_Chi_Minh)"
    )
    print("=" * 65)

    # Khởi tạo kết nối Google Sheets
    spreadsheet = None
    sheet_names = []
    try:
        client = get_gspread_client()
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        sheet_names = [s.title for s in spreadsheet.worksheets()]
        print(f"✅ Kết nối Google Sheets thành công! Spreadsheet ID: {SPREADSHEET_ID}")
    except Exception as e:
        print(f"⚠️ Cảnh báo kết nối Google Sheets: {e}")

    driver = None
    for attempt_login in range(5):
        driver = login()
        if driver:
            print("✅ Đăng nhập Teams thành công!")
            break
        else:
            print(f"⚠️ Thử đăng nhập lại lần {attempt_login + 1}/5...")
            time.sleep(2)

    if not driver:
        print("❌ Đăng nhập Teams không thành công!")
        exit(1)

    has_error = False
    try:
        driver.save_screenshot("after_login.png")

        # Thử mở nhóm chat với retry
        chat_opened = False
        for attempt in range(1, 4):
            print(f"📂 Đang mở nhóm chat '{chat}' (lần thử {attempt}/3)...")
            if open_chat(driver, chat):
                chat_opened = True
                break
            time.sleep(3)

        if not chat_opened:
            print(f"❌ KHÔNG thể mở nhóm chat '{chat}' trên Teams sau 3 lần thử!")

        if spreadsheet and sheet_names:
            messages = get_filtered_messages(
                spreadsheet,
                sheet_names,
                time_window=(start_dt, end_dt),
            )
            combined_msgs = combine_messages(messages)

            print(
                f"\n✅ Báo cáo lọc được cho [{shift_name}] "
                f"({start_dt.strftime('%Y-%m-%d %H:%M:%S')} -> {end_dt.strftime('%Y-%m-%d %H:%M:%S')}):"
            )

            if combined_msgs and chat_opened:
                for sheet_name, msg_content in combined_msgs.items():
                    print(
                        f"Testing\nSheet: [ {sheet_name} ]\nMessage: [ {msg_content} ]\n"
                    )
                    message = f"[ {sheet_name} ]\n" + msg_content
                    send_message(driver, message)
            elif combined_msgs and not chat_opened:
                print(
                    f"⚠️ Có {len(combined_msgs)} báo cáo cần gửi nhưng bỏ qua gửi Teams do không mở được nhóm chat '{chat}'."
                )
                has_error = True
            else:
                print("ℹ️ Không có tin nhắn báo cáo nào trong khung giờ này.")

            write_to_sheet(spreadsheet, sheet_names, "Report", messages)
            write_to_sheet(spreadsheet, sheet_names, "GetReport", messages)
        else:
            print("⚠️ Bỏ qua lọc và ghi báo cáo do không có kết nối Google Sheets.")
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
        print("✅ Hoàn tất toàn bộ công việc!")
        if has_error:
            raise RuntimeError(
                f"Gặp lỗi khi mở chat '{chat}' trên Teams để gửi báo cáo."
            )


if __name__ == "__main__":
    run_report()
