import os
import gc
import re
import time
import datetime
import pytz
from dotenv import load_dotenv, find_dotenv

# Tự động tìm và nạp file .env ở thư mục hiện tại hoặc thư mục cha
load_dotenv(find_dotenv())

# Tắt GC để tránh lỗi undetected_chromedriver khi thu hồi handle
gc.disable()

# =========================
# CONFIG
# =========================
os.environ["TZ"] = "Asia/Ho_Chi_Minh"
if hasattr(time, "tzset"):
    time.tzset()

email = os.environ.get("TEAMS_EMAIL")
password = os.environ.get("TEAMS_PASSWORD")
chat = os.environ.get("TEAMS_CHAT", "GetReport")
local_tz = pytz.timezone("Asia/Ho_Chi_Minh")
SPREADSHEET_ID = os.environ.get(
    "SPREADSHEET_ID", "1_m7s-1-I-SOFfzlWe7CBf5fstFir7qXYAKW4j-8hKYM"
)
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

MESSAGE_PATTERN = re.compile(
    r".*(?:[+-]\s*\d+\s*/|check\s*[-_ ]?out).*",
    re.IGNORECASE | re.DOTALL,
)
CHECKOUT_PATTERN = re.compile(r"check\s*[-_ ]?out", re.IGNORECASE)
CHECKIN_PATTERN = re.compile(r"check\s*[-_ ]?in", re.IGNORECASE)
RESET_PATTERN = re.compile(r"^\s*reset\s*15\s*min", re.IGNORECASE)


# =========================
# TIME WINDOW CHO GET_REPORT (2 CA CRONJOB: 08:00 AM VÀ 02:00 PM / 14:00 - GIỜ HỒ CHÍ MINH)
# =========================
def get_hcm_now():
    """Luôn lấy giờ hiện tại chuẩn theo múi giờ Asia/Ho_Chi_Minh (UTC+7)."""
    return datetime.datetime.now(datetime.timezone.utc).astimezone(local_tz)


def get_report_time_window(now=None):
    """
    Đồng bộ khung giờ lọc báo cáo với 3 ca cào tin nhắn của get_message (02:00, 13:00, 18:00):
    - Ca 08:00 AM (chạy buổi sáng, hour < 12):
        Lấy tin nhắn từ 13:00:00 hôm qua -> 02:00:00 sáng hôm nay
        (Tổng hợp dữ liệu của 2 ca get_message lúc 18:00 hôm qua và 02:00 sáng nay).
    - Ca 02:00 PM / 14:00 (chạy buổi chiều, hour >= 12):
        Lấy tin nhắn từ 02:00:00 sáng hôm nay -> 14:00:00 chiều hôm nay
        (Tổng hợp dữ liệu của ca get_message lúc 13:00 trưa nay).
    """
    if now is None:
        now = get_hcm_now()
    elif now.tzinfo is None:
        now = local_tz.localize(now)
    else:
        now = now.astimezone(local_tz)

    ref_time = now
    ci_trigger_str = os.getenv("CI_TRIGGER_TIME", "").strip()
    if ci_trigger_str:
        try:
            parsed_trigger = datetime.datetime.strptime(
                ci_trigger_str, "%Y-%m-%d %H:%M:%S"
            )
            ref_time = local_tz.localize(parsed_trigger)
        except Exception:
            ref_time = now

    ref_date = ref_time.date()
    ref_hour = ref_time.hour

    if ref_hour < 12:
        shift_name = "Ca báo cáo 08:00 AM (8h sáng)"
        start_dt = datetime.datetime.combine(
            ref_date - datetime.timedelta(days=1),
            datetime.time(13, 0, 0),
        )
        end_dt = datetime.datetime.combine(
            ref_date,
            datetime.time(2, 0, 0),
        )
    else:
        shift_name = "Ca báo cáo 02:00 PM (14h chiều)"
        start_dt = datetime.datetime.combine(
            ref_date,
            datetime.time(2, 0, 0),
        )
        end_dt = datetime.datetime.combine(
            ref_date,
            datetime.time(14, 0, 0),
        )

    return start_dt, end_dt, shift_name
