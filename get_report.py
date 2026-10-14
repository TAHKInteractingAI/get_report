import os
import re
import time
import json
import datetime
import pytz
from dateutil import parser
from dotenv import load_dotenv, find_dotenv

import gspread
from google.auth import default
from google.oauth2.service_account import Credentials

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import undetected_chromedriver as uc

import platform
import subprocess
import re


# Tự động tìm và nạp file .env ở thư mục hiện tại hoặc thư mục cha
load_dotenv(find_dotenv())

# =========================
# PATCH UNDETECTED CHROMEDRIVER CLEANUP
# =========================
# Sửa lỗi OSError: [WinError 6] The handle is invalid khi Python giải phóng bộ nhớ (garbage collection)
def _patch_uc_del():
    def _safe_del(self):
        try:
            self.quit()
        except Exception:
            pass
    uc.Chrome.__del__ = _safe_del

_patch_uc_del()

# =========================
# CONFIG
# =========================
os.environ["TZ"] = "Asia/Ho_Chi_Minh"
if hasattr(time, "tzset"):
    time.tzset()

email = os.environ.get("TEAMS_EMAIL")
password = os.environ.get("TEAMS_PASSWORD")
chat = "GetReport"
local_tz = pytz.timezone("Asia/Ho_Chi_Minh")
SPREADSHEET_ID = "1_m7s-1-I-SOFfzlWe7CBf5fstFir7qXYAKW4j-8hKYM"
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
# GOOGLE SHEETS AUTHENTICATION
# =========================
def get_gspread_client(credentials_path="gcp-credentials.json"):
    """Khởi tạo và xác thực kết nối với Google Sheets API linh hoạt từ nhiều nguồn."""
    # 1. Ưu tiên đọc từ các biến môi trường dạng JSON trong file .env
    env_json = (
        os.getenv("GCP_SA_KEY")
        or os.getenv("GCP_CREDENTIALS_JSON")
        or os.getenv("GOOGLE_CREDENTIALS")
    )
    if env_json and env_json.strip().startswith("{"):
        try:
            info = json.loads(env_json)
            creds = Credentials.from_service_account_info(info, scopes=SCOPES)
            return gspread.authorize(creds)
        except Exception as e:
            print(f"⚠️ Lỗi đọc credentials từ biến môi trường: {e}")

    # 2. Đọc từ file JSON trên ổ đĩa (thử nhiều đường dẫn bao gồm file gcp-credentials.json ở thư mục get_message)
    possible_paths = [
        credentials_path,
        credentials_path + ".json",
        "gcp-credentials.json",
        "gcp-credentials.json.json",
        "credentials.json",
        os.path.join("..", "get_message", "gcp-credentials.json"),
        os.path.join("..", "get_message", "gcp-credentials.json.json"),
    ]

    for path in possible_paths:
        if os.path.exists(path):
            try:
                creds = Credentials.from_service_account_file(path, scopes=SCOPES)
                return gspread.authorize(creds)
            except Exception as e:
                print(f"⚠️ Lỗi đọc file credentials '{path}': {e}")

    # 3. Thử dùng default credentials
    try:
        creds, _ = default()
        return gspread.authorize(creds)
    except Exception as e:
        raise FileNotFoundError(
            "Không thể khởi tạo kết nối Google Sheets API. "
            "Vui lòng cấu hình biến GCP_SA_KEY / GCP_CREDENTIALS_JSON trong file .env hoặc thêm file 'gcp-credentials.json' vào thư mục."
        ) from e


def display_screenshot(driver: webdriver.Chrome, file_name: str = "screenshot.png"):
    """Chụp màn hình và hiển thị"""
    try:
        driver.save_screenshot(file_name)
        print(f"📸 Đã lưu ảnh màn hình '{file_name}'")
    except Exception as e:
        print(f"❌ Không thể lưu ảnh màn hình '{file_name}': {e}")
    time.sleep(3)


save_screenshot = display_screenshot


def send_message(driver, message):
    try:
        message_box = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"]'))
        )

        for line in message.split("\n"):
            try:
                message_box.send_keys(line)
            except Exception:
                # Dự phòng khi dòng tin nhắn có ký tự ngoài BMP (Emoji) khiến ChromeDriver send_keys báo lỗi
                safe_line = "".join(c for c in line if ord(c) <= 0xFFFF)
                message_box.send_keys(safe_line)
            message_box.send_keys(Keys.SHIFT, Keys.ENTER)

        display_screenshot(driver, "after_typing_message.png")
        time.sleep(3)
        message_box.send_keys(Keys.ENTER)
        time.sleep(3)
        display_screenshot(driver, "after_sending_message.png")

    except Exception as e:
        print(f"❌ Lỗi khi gửi tin nhắn: {e}")


def open_chat(driver, chat_name):
    """Mở nhóm chat an toàn bằng cả danh sách ngoài và thanh Filter bên trái."""
    wait = WebDriverWait(driver, 20)
    chat_item_xpath = (
        '//*[contains(@data-tid, "chat-list-item") '
        'or contains(@data-tid, "chat-item") '
        'or @data-item-type="chat" '
        'or @role="listitem" '
        'or @role="treeitem"]'
    )

    try:
        # 1. Thử tìm trực tiếp trên danh sách hiển thị
        wait.until(EC.presence_of_element_located((By.XPATH, chat_item_xpath)))
        groups = driver.find_elements(By.XPATH, chat_item_xpath)

        for g in groups:
            lines = [x.strip() for x in g.text.splitlines() if x.strip()]
            if not lines:
                continue
            if lines[0] in ["Unread", "Chưa đọc"]:
                lines.pop(0)
            if not lines:
                continue

            txt = lines[0]
            if not txt:
                txt = g.get_attribute("aria-label") or ""

            if txt == chat_name or chat_name.startswith(txt.replace("...", "").strip()):
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", g)
                time.sleep(1)
                g.click()
                time.sleep(3)
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"]'))
                )
                display_screenshot(driver, "after_opening_chat.png")
                print(f"📂 Đã mở đúng nhóm: {chat_name}")
                return True

        print(f"⚠️ Không thấy {chat_name} ở ngoài, thử dùng thanh Filter...")

        # 2. Mở thanh Filter bên trái nếu đang ẩn
        try:
            filter_icon_xpath = (
                '//button[@data-testid="simple-collab-left-rail-header-sticky-filter-v2-button"]'
                ' | //button[contains(@aria-keyshortcuts, "Ctrl+Shift+F")]'
            )
            filter_icon = driver.find_element(By.XPATH, filter_icon_xpath)
            driver.execute_script("arguments[0].click();", filter_icon)
            time.sleep(1)
        except Exception:
            pass

        left_search_xpath = (
            '//input[@data-testid="simple-collab-left-rail-sticky-filter-input"]'
            ' | //input[@id="simple-collab-left-rail-sticky-filter-input-id"]'
            ' | //input[contains(@placeholder, "Filter") or contains(@placeholder, "Lọc")]'
        )
        search = wait.until(EC.presence_of_element_located((By.XPATH, left_search_xpath)))
        search.click()
        search.send_keys(Keys.CONTROL + "a")
        search.send_keys(Keys.BACKSPACE)
        search.send_keys(chat_name)
        time.sleep(3)

        filtered_result_xpath = f"//span[normalize-space(text())='{chat_name}']"
        filtered_result = wait.until(
            EC.presence_of_element_located((By.XPATH, filtered_result_xpath))
        )
        driver.execute_script("arguments[0].click();", filtered_result)
        time.sleep(3)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"]'))
        )
        display_screenshot(driver, "after_opening_chat.png")
        print(f"📂 Đã mở nhóm qua thanh Filter: {chat_name}")
        return True

    except Exception as e:
        display_screenshot(driver, "open_chat_error.png")
        print(f"❌ Lỗi khi mở chat '{chat_name}': {e}")
        return False


def combine_messages(messages_dict):
    combined = {}
    for sheet_name, msg_list in messages_dict.items():
        if msg_list:
            combined[sheet_name] = "\n\n".join(msg_list)
    return combined


def preprocess_message(content):
    content = str(content or "").replace("\xa0", " ")
    content = re.sub(r"-\s+-", "-", content)
    # Tách các trường tiêu đề / báo cáo bị dính liền dòng (ưu tiên cụm dài trước)
    content = re.sub(
        r"(?<!\n)\s*((?:tiêu\s*đề\s*task|task|kết\s*quả(?:\s*làm\s*được)?|tiến\s*độ(?:\s*công\s*việc)?|link\s*proof)\s*:)",
        r"\n\1",
        content,
        flags=re.IGNORECASE,
    )
    # Tách mục 6/..10/ bị dính liền chữ phía trước (ví dụ: 'bot- 6/', '...hôm nay+ 6/')
    content = re.sub(r"([^\n\s])\s*([+-]\s*(?:[6-9]|10)\s*/)", r"\1\n\2", content)
    content = re.sub(r"\s*[+]+\s*(\d+\s*/)\s*", r"\n+ \1 ", content)
    content = re.sub(
        r"\s*[-]+\s*((?:6|7|8|9|10)\s*/)\s*",
        r"\n- \1 ",
        content,
    )

    lines = content.splitlines()
    processed_lines = []

    for line in lines:
        line = line.strip()
        if not line:
            continue

        if re.match(r"^[+*•-]?\s*(?:6|7|8|9|10)\s*/", line):
            continue

        if re.match(r"^(\+|\d+\.|=>|-)", line):
            line = "\u200b" + line
        processed_lines.append(line)

    changed = True
    while changed:
        changed = False
        n = len(processed_lines)
        for L in range(n // 2, 0, -1):
            for i in range(n - 2 * L + 1):
                if processed_lines[i : i + L] == processed_lines[i + L : i + 2 * L]:
                    processed_lines = (
                        processed_lines[: i + L] + processed_lines[i + 2 * L :]
                    )
                    changed = True
                    break
            if changed:
                break

    content = "\n".join(processed_lines)
    return content.strip()


def is_valid_message(content):
    if not content or not str(content).strip():
        return False
    text = str(content).strip()
    if RESET_PATTERN.search(text):
        return False

    # Dấu hiệu chắc chắn là báo cáo kết quả / hoàn thành ca (CheckOut / Report):
    # 1. Có từ khóa checkout / check-out / check out
    has_checkout_kw = bool(CHECKOUT_PATTERN.search(text))
    # 2. Có checklist items 6..10
    has_checklist_items = bool(re.search(r"[+-]\s*(?:[6-9]|10)\s*/", text))
    # 3. Có báo cáo kết quả công việc / tiêu đề task hoàn thành
    has_work_results = bool(
        re.search(r"(?:tiêu\s*đề\s*task|kết\s*quả|tiến\s*độ|link\s*proof)", text, re.IGNORECASE)
    )

    is_report = has_checkout_kw or has_checklist_items or has_work_results

    first_line = next((ln.strip() for ln in text.splitlines() if ln.strip()), "")
    is_checkin_header = bool(CHECKIN_PATTERN.search(first_line))

    # Nếu dòng đầu là CheckIn nhưng KHÔNG CÓ bất kỳ kết quả, tiêu đề task hay checklist 6-10 -> CheckIn thật
    if is_checkin_header and not (has_checklist_items or has_work_results or has_checkout_kw):
        return False

    return is_report


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
            parsed_trigger = datetime.datetime.strptime(ci_trigger_str, "%Y-%m-%d %H:%M:%S")
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


def _parse_sheet_row_datetime(date_val, time_val):
    """Đọc DATE (YYYY-MM-DD) và TIME (HH:MM:SS) từ Google Sheet một cách chính xác, tránh đảo ngày/tháng."""
    date_str = str(date_val).strip()
    time_str = str(time_val).strip()
    if not date_str or not time_str:
        return None

    date_part = None
    for d_fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            date_part = datetime.datetime.strptime(date_str, d_fmt).date()
            break
        except ValueError:
            continue
    if date_part is None:
        date_part = parser.parse(date_str).date()

    time_part = None
    for t_fmt in ("%H:%M:%S", "%H:%M"):
        try:
            time_part = datetime.datetime.strptime(time_str, t_fmt).time()
            break
        except ValueError:
            continue
    if time_part is None:
        time_part = parser.parse(time_str).time()

    return datetime.datetime.combine(date_part, time_part)


def get_filtered_messages(spreadsheet, sheet_names, current_hour=None, time_window=None):
    if time_window is not None:
        start_dt, end_dt = time_window
    else:
        start_dt, end_dt, _ = get_report_time_window()

    messages = {sheet_name: [] for sheet_name in sheet_names}
    EXCLUDED_SHEETS = [
        "Report",
        "GetReport",
        "iX000s iSSale TTS Base.XoắnNỆN50k*CấuTrúcVolunt",
        "iX000s iSSale gbBOSS AH*AU*cOL*YeuCauTop-iUp*KTra",
    ]

    for sheet_name in sheet_names:
        if sheet_name in EXCLUDED_SHEETS:
            continue
        try:
            sheet = spreadsheet.worksheet(sheet_name)
            data = sheet.get_all_records()

            for row in data:
                try:
                    full_datetime = _parse_sheet_row_datetime(row.get("DATE", ""), row.get("TIME", ""))
                    if full_datetime is None:
                        continue

                    raw_content = str(row.get("CONTENT", "")).strip()

                    if not is_valid_message(raw_content):
                        continue

                    content = preprocess_message(raw_content)
                    if not content:
                        continue

                    is_in_time = start_dt <= full_datetime <= end_dt

                    if is_in_time and content not in messages[sheet_name]:
                        messages[sheet_name].append(content)
                except Exception:
                    continue
        except Exception as e:
            print(f"❌ Lỗi sheet '{sheet_name}': {e}")
    return messages


def write_to_sheet(spreadsheet, sheet_names, sheet_target_name, messages):
    try:
        sheet_names_with_data = [name for name in sheet_names if messages.get(name)]
        if not sheet_names_with_data:
            print(f"--- Không có dữ liệu để ghi vào {sheet_target_name} ---")
            return

        try:
            ws = spreadsheet.worksheet(sheet_target_name)
        except gspread.exceptions.WorksheetNotFound:
            ws = spreadsheet.add_worksheet(
                title=sheet_target_name, rows="1000", cols="20"
            )

        existing_data = ws.get_all_values()
        max_len = max(len(messages[s]) for s in sheet_names_with_data)

        rows_to_append = []
        for i in range(max_len):
            row = []
            for sheet_name in sheet_names_with_data:
                msg = messages[sheet_name][i] if i < len(messages[sheet_name]) else ""
                row.append(msg)

            if row not in existing_data:
                rows_to_append.append(row)

        if rows_to_append:
            ws.append_rows(rows_to_append, value_input_option="USER_ENTERED")
            print(
                f"✅ Đã ghi thêm {len(rows_to_append)} dòng mới vào [{sheet_target_name}]"
            )
        else:
            print(f"ℹ️ Không có dữ liệu mới (trùng lặp) cho [{sheet_target_name}]")

    except Exception as e:
        print(f"❌ Lỗi khi ghi vào sheet {sheet_target_name}: {e}")
# =========================
# Kiểm tra version Chrome
# =========================       
def get_installed_chrome_major_version():
    """Tự động kiểm tra Major Version của Chrome trên máy"""

    system = platform.system()
    try:
        if system == "Windows":
            import winreg
            # Đọc phiên bản Chrome từ Registry Windows
            try:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon")
            except FileNotFoundError:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon")
            version, _ = winreg.QueryValueEx(key, "version")
            return int(version.split('.')[0])

        elif system == "Linux":
            output = subprocess.check_output(["google-chrome", "--version"]).decode("utf-8")
            match = re.search(r"Google Chrome (\d+)\.", output)
            if match:
                return int(match.group(1))

        elif system == "Darwin":  # macOS
            cmd = r"/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --version"
            output = subprocess.check_output(cmd, shell=True).decode("utf-8")
            match = re.search(r"Google Chrome (\d+)\.", output)
            if match:
                return int(match.group(1))
    except Exception as e:
        print(f"⚠️ Không thể tự động phát hiện phiên bản Chrome: {e}")
    
    return None



def get_driver():
    options = uc.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.page_load_strategy = "eager"
    options.add_argument("--lang=en-GB")
    
    prefs = {
        "profile.cookie_controls_mode": 0,
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False
    } 

    options.add_experimental_option("prefs", prefs)

    proxy_url = os.getenv("PROXY_URL")
    if proxy_url:
        options.add_argument(f"--proxy-server={proxy_url}")

    import subprocess
    import re

    chrome_version = None
    try:
        result = subprocess.check_output(["google-chrome", "--version"]).decode("utf-8")
        chrome_version = int(re.search(r"\d+", result).group(0))
        print(
            f"✅ Đã tự động nhận diện Chrome trên máy chủ là version: {chrome_version}"
        )
    except Exception:
        chrome_version = get_installed_chrome_major_version()

    if chrome_version:
        driver = uc.Chrome(options=options, version_main=chrome_version)
    else:
        driver = uc.Chrome(options=options)

    try:
        driver.execute_cdp_cmd(
            "Emulation.setTimezoneOverride",
            {"timezoneId": "Asia/Ho_Chi_Minh"},
        )
    except Exception:
        pass

    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {"source": """
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            window.navigator.chrome = { runtime: {} };
            Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3, 4, 5] });
            Object.defineProperty(navigator, 'languages', { get: () => ['en-GB', 'en-US', 'en'] });
            Object.defineProperty(navigator, 'credentials', {
                get: () => undefined
            });

            window.PublicKeyCredential = undefined;
        """},
    )

    return driver


def login():
    driver = get_driver()

    # Truy cập link chuẩn cho Work/School
    driver.get("https://teams.microsoft.com/")
    wait = WebDriverWait(driver, 30)

    try:
        print("⏳ Đang đăng nhập...")

        # 1. Xử lý nút Sign in (nếu bị đẩy ra trang chờ)
        try:
            sign_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        '//button[contains(., "Sign in")] | //a[contains(., "Sign in")] | //button[contains(., "Đăng nhập")]',
                    )
                )
            )
            sign_btn.click()
        except:
            pass  # Bỏ qua nếu form điền email hiện ra trực tiếp

        # 2. Ô nhập Email (Sử dụng Selector linh hoạt cho Microsoft)
        email_box = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'input[type="email"], input[name="loginfmt"]')
            )
        )
        email_box.send_keys(email)
        email_box.send_keys(Keys.RETURN)

        time.sleep(3)
        # ====== THÊM ĐOẠN NÀY VÀO ======
        # Xử lý trường hợp Microsoft đòi gửi mã code, ép nó quay về dùng Mật khẩu
        try:
            use_pass_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        '//*[contains(text(), "Use your password") or contains(text(), "Sử dụng mật khẩu")]',
                    )
                )
            )
            use_pass_btn.click()
            time.sleep(2)
        except:
            pass  # Nếu màn hình đi thẳng tới ô mật khẩu thì cứ bỏ qua bước này
        # ===============================
        # ====== XỬ LÝ MÀN HÌNH "Sign in another way" -> CHỌN USE YOUR PASSWORD ======
        try:
            # 1. Nếu có màn hình "Other ways to sign in" thì bấm trước
            try:
                other_ways_btn = WebDriverWait(driver, 4).until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            '//*[contains(text(), "Other ways to sign in") or contains(text(), "Cách đăng nhập khác")]',
                        )
                    )
                )
                driver.execute_script("arguments[0].click();", other_ways_btn)
                print("👉 Đã chọn: Other ways to sign in")
                time.sleep(2)
            except:
                pass

            # 2. Chọn dòng "Use your password" (Dùng JS click để không bị trượt)
            use_pass_btn = WebDriverWait(driver, 6).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        '//*[contains(text(), "Use your password") or contains(text(), "Sử dụng mật khẩu")]/ancestor::div[@role="button"] '
                        '| //*[contains(text(), "Use your password") or contains(text(), "Sử dụng mật khẩu")]',
                    )
                )
            )
            driver.execute_script("arguments[0].click();", use_pass_btn)
            print("👉 Đã kích hoạt: Use your password")
            time.sleep(3)
        except Exception as e:
            print("ℹ️ Bỏ qua chọn phương thức (hoặc đã ở màn hình nhập pass):", e)
        # ===========================================================================
        # 3. Ô nhập Password
        pass_box = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'input[type="password"], input[name="passwd"]')
            )
        )
        pass_box.send_keys(password)
        time.sleep(1)
        pass_box.send_keys(Keys.RETURN)
        time.sleep(3)

        # 4. Xử lý nút "Stay signed in?" (Chọn No để không lưu đăng nhập)
        try:
            print("⏳ Đang xử lý màn hình Stay signed in...")
            no_btn = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        '//*[@id="declineButton"] | //*[@id="idBtn_Back"] | //*[@value="No"] '
                        '| //button[@data-testid="secondaryButton"] '
                        '| //button[normalize-space(.)="No" or normalize-space(.)="Không"]',
                    )
                )
            )
            no_btn.click()
            time.sleep(3)
        except:
            print("⚠️ Không thấy màn hình Stay signed in, tiếp tục...")
            pass

        try:
            body_text = driver.find_element(By.TAG_NAME, "body").text.lower()
            if "try again later" in body_text or "can't sign you in right now" in body_text:
                save_screenshot(driver, "try_again_later_blocked.png")
                raise RuntimeError(
                    "Microsoft tạm khóa phiên đăng nhập (Try again later / We can't sign you in right now)."
                )
        except RuntimeError:
            raise
        except Exception:
            pass

        print("✅ Đăng nhập thành công")

        # Chờ giao diện Teams load hẳn
        time.sleep(15)

        return driver

    except Exception as e:
        save_screenshot(driver, "login_error.png")
        print("❌ Đăng nhập thất bại:", e)
        try:
            driver.quit()
        except Exception:
            pass
        return None


if __name__ == "__main__":
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

    try:
        driver.save_screenshot("after_login.png")
        open_chat(driver, chat)

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

            for sheet_name, msg_content in combined_msgs.items():
                print(f"Testing\nSheet: [ {sheet_name} ]\nMessage: [ {msg_content} ]\n")
                message = f"[ {sheet_name} ]\n" + msg_content
                send_message(driver, message)

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
