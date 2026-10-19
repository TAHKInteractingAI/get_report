import os
import re
import json
import datetime
from dateutil import parser
import gspread
from google.auth import default
from google.oauth2.service_account import Credentials
from .config import (
    SCOPES,
    MESSAGE_PATTERN,
    RESULT_PATTERN,
    CHECKOUT_PATTERN,
    CHECKIN_PATTERN,
    RESET_PATTERN,
    get_report_time_window,
)


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

    # 2. Đọc từ file JSON trên ổ đĩa
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


def combine_messages(messages_dict):
    combined = {}
    for sheet_name, msg_list in messages_dict.items():
        if msg_list:
            combined[sheet_name] = "\n\n".join(msg_list)
    return combined


def preprocess_message(content):
    content = str(content or "").replace("\xa0", " ")
    content = re.sub(r"-\s+-", "-", content)
    # Tách các trường báo cáo bị dính liền dòng
    content = re.sub(
        r"(?<!\n)\s*((?:task|kết\s*quả|tiến\s*độ(?:\s*công\s*việc)?|link\s*proof)\s*:)",
        r"\n\1",
        content,
        flags=re.IGNORECASE,
    )
    content = re.sub(r"\s*[+]+\s*(\d+\s*/)\s*", r"\n+ \1 ", content)
    content = re.sub(
        r"(?:^|\n|\s)[-]+\s*((?:6|7|8|9|10)\s*/)\s*",
        r"\n- \1 ",
        content,
    )

    lines = content.splitlines()
    processed_lines = []

    for line in lines:
        line = line.strip()
        if not line:
            continue

        if re.match(r"^[+-]?\s*(6|7|8|9|10)\s*/", line):
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

    has_checkout = bool(CHECKOUT_PATTERN.search(text))
    has_result = bool(RESULT_PATTERN.search(text))

    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    header = " ".join(lines[:2])
    has_checkin = bool(CHECKIN_PATTERN.search(header))

    # 1. Có từ khóa checkout -> Chắc chắn là báo cáo kết ca hợp lệ
    if has_checkout:
        return True

    # 2. Nếu ở tiêu đề có check in và KHÔNG có checkout:
    # Bắt buộc phải có phần KẾT QUẢ/TIẾN ĐỘ cụ thể (để cứu trường hợp gõ nhầm check in thay vì check out)
    # Nếu không có kết quả/tiến độ -> Chắc chắn là tin nhắn CHECK IN điểm danh đầu ca -> BỎ QUA
    if has_checkin:
        return has_result

    # 3. Không có cả check in lẫn check out: chỉ chấp nhận nếu có báo cáo kết quả hoặc match pattern báo cáo
    return has_result or bool(MESSAGE_PATTERN.match(text))


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


def get_filtered_messages(
    spreadsheet, sheet_names, current_hour=None, time_window=None
):
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
                    full_datetime = _parse_sheet_row_datetime(
                        row.get("DATE", ""), row.get("TIME", "")
                    )
                    if full_datetime is None:
                        continue

                    raw_content = str(row.get("CONTENT", "")).strip()
                    if not raw_content:
                        continue

                    # Tiền xử lý để tách dòng các trường bị dính liền (task:, kết quả:, link proof:,...)
                    content = preprocess_message(raw_content)
                    if not content:
                        continue

                    if not is_valid_message(content):
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
