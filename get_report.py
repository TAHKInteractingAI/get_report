import datetime
import pytz
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from dateutil import parser
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import gc
gc.disable() # 03/23/2022 added this one to prevent trash collection and avoide crashing the notebooks
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from IPython.display import Image, display
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
import time
import json
import os

email = os.environ.get('TEAMS_EMAIL')
password = os.environ.get('TEAMS_PASSWORD')
gcp_credentials_json = os.environ.get('GCP_SA_KEY')
chat = "GetReport"
local_tz = pytz.timezone("Asia/Ho_Chi_Minh")
SPREADSHEET_ID = '1_m7s-1-I-SOFfzlWe7CBf5fstFir7qXYAKW4j-8hKYM'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# Khởi tạo kết nối
creds_dict = json.loads(gcp_credentials_json)
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, SCOPES)
client = gspread.authorize(creds)
spreadsheet = client.open_by_key(SPREADSHEET_ID)

# Lấy danh sách tất cả sheet
sheet_names = [s.title for s in spreadsheet.worksheets()]

MESSAGE_PATTERN = re.compile(
    r".*\+\s*([6-9])/.*",
    re.IGNORECASE | re.DOTALL
)

def display_screenshot(driver: webdriver.Chrome, file_name: str = 'screenshot.png'):
    """Chụp màn hình và hiển thị"""
    # driver.execute_script("window.scrollBy(0, 250);")
    driver.save_screenshot(file_name)
    time.sleep(3)
    display(Image(filename=file_name))

def send_message(driver, message):
    try:
        message_box = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"]'))
        )

        # Gửi từng dòng, xuống dòng bằng Shift + Enter
        for line in message.split('\n'):
            message_box.send_keys(line)
            message_box.send_keys(Keys.SHIFT, Keys.ENTER)
            display_screenshot(driver)

        # Gửi tin nhắn (ENTER)
        message_box.send_keys(Keys.ENTER)
        #time.sleep(2)
        display_screenshot(driver)
        time.sleep(2)  # Giữ lại chút delay sau khi gửi

    except Exception as e:
        print(f"❌ Lỗi khi gửi tin nhắn: {e}")

def login():
    import tempfile
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    temp_dir = tempfile.mkdtemp()
    options.add_argument(f"--user-data-dir={temp_dir}")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get("https://teams.live.com/v2/")
    time.sleep(5)

    sign_in_btn = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//button[@type="button" and contains(., "Sign in")]'))
    )
    sign_in_btn.click()

    email_input = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "usernameEntry")))
    email_input.send_keys(email)
    email_input.send_keys(Keys.RETURN)

    # Chọn 'Use your password' nếu xuất hiện
    try:
        use_pass_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//span[@role="button" and contains(text(), "Use your password")]'))
        )
        use_pass_btn.click()
    except Exception as e:
        print("Không tìm thấy nút 'Use your password'.")

    # Tiếp tục nhập mật khẩu như cũ
    password_input = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "passwordEntry")))
    password_input.send_keys(password)
    password_input.send_keys(Keys.RETURN)

    # driver.save_screenshot("after_email.png")
    # print("Đã chụp màn hình sau khi nhập email.")

    try:
        no_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="secondaryButton"]'))
        )
        no_button.click()
    except Exception as e:
        print("Không tìm thấy nút 'No'.")


    try:
        button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="primaryButton"]'))
        )
        button.click()
    except:
        pass
    
    return driver

def open_chat(driver, chat_name):
    try:
        chat_element = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, f"//span[contains(text(), '{chat_name}')]"))
        )
        chat_element.click()
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"]'))
        )
        display_screenshot(driver)
    except Exception as e:
        print(f"❌ Lỗi khi mở chat '{chat_name}': {e}")

def combine_messages(messages_dict):
    """Gộp các message trong cùng sheet thành một chuỗi"""
    combined = {}
    for sheet_name, msg_list in messages_dict.items():
        if msg_list:
            # Gộp các message cách nhau bởi 2 dòng mới
            combined[sheet_name] = "\n\n".join(msg_list)
    return combined

def preprocess_message(content):
    # Bước 1: Chuẩn hóa các pattern đặc biệt
    content = re.sub(r"-\s+-", "-", content)
    content = re.sub(r"\s*[+]+\s*(\d+/)\s*", r"\n+ \1 ", content)
    content = re.sub(
        r"(\+\s*\d+/.*?)\n(?![+\d]|Link|http)",  # Tìm xuống dòng KHÔNG bắt đầu bằng +, số, •, hoặc link
        r"\1 ",  # Thay thế xuống dòng bằng khoảng trắng
        content,
        flags=re.DOTALL
    )

    # Bước 2: Xử lý xuống dòng và thêm ZWSP
    lines = content.splitlines()
    processed_lines = []

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Thêm ZWSP nếu dòng bắt đầu bằng +, =>, hoặc số
        if re.match(r"^(\+|\d+\.|=>|-)", line):
            line = "\u200B" + line  # Chèn Zero Width Space ở đầu dòng

        processed_lines.append(line)

    # Bước 3: Gộp lại và chuẩn hóa
    content = "\n".join(processed_lines)
    content = re.sub(r"\n{3,}", "\n\n", content)  # Giảm nhiều xuống dòng
    return content

def is_valid_message(content):
    """Kiểm tra message có đúng định dạng"""
    return bool(MESSAGE_PATTERN.match(content))

# Hàm lọc theo thời gian
def get_filtered_messages(current_hour):
    tz = pytz.timezone('Asia/Ho_Chi_Minh')
    now = datetime.datetime.now(tz)

    messages = {sheet_name: [] for sheet_name in sheet_names}  # Khởi tạo dictionary để lưu message theo sheet name
    EXCLUDED_SHEETS = ["iX000s iSSale TTS Base.XoắnNỆN50k*CấuTrúcVolunt", "iX000s iSSale gbBOSS AH*AU*cOL*YeuCauTop-iUp*KTra"]

    for sheet_name in sheet_names:
        try:
            if sheet_name in EXCLUDED_SHEETS:
                continue
            sheet = spreadsheet.worksheet(sheet_name)
            data = sheet.get_all_records()

            for row in data:
                try:
                    date_str = str(row['DATE'])  # Giả định định dạng YYYY-MM-DD
                    time_str = str(row['TIME'])  # Giả định định dạng HH:MM hoặc HH:MM:SS

                    date_obj = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
                    time_obj = parser.parse(time_str).time()

                    full_datetime = datetime.datetime.combine(date_obj, time_obj)

                    content = row['CONTENT'].strip()
                    content = preprocess_message(content)

                    # Bỏ qua message không hợp lệ
                    if not is_valid_message(content):
                        continue

                    if current_hour == 8:
                        # Lọc từ 13h hôm qua đến 1h sáng hôm nay
                        start_time = datetime.datetime.combine(now.date() - datetime.timedelta(days=1), datetime.time(13, 0))  # 13h hôm qua
                        end_time = datetime.datetime.combine(now.date(), datetime.time(1, 0))  # 1h sáng hôm nay
                        if start_time <= full_datetime < end_time:
                            messages[sheet_name].append(content)

                    elif current_hour == 14:
                        # Lọc từ 1h sáng đến 13h hôm nay
                        start_time = datetime.datetime.combine(now.date(), datetime.time(1, 0))  # 1h sáng hôm nay
                        end_time = datetime.datetime.combine(now.date(), datetime.time(13, 0))  # 13h hôm nay
                        if start_time <= full_datetime < end_time:
                            messages[sheet_name].append(content)

                except:
                    continue
        except Exception as e:
            print(f"❌ Lỗi truy cập sheet '{sheet_name}': {e}")

    return messages

# Tạo sheet mới để lưu kết quả
def create_or_update_report_sheet(messages):
    try:
        # Lọc những sheet có message
        sheet_names_with_data = [sheet_name for sheet_name in sheet_names if messages[sheet_name]]

        # Kiểm tra xem sheet "Report" đã tồn tại chưa
        try:
            report_sheet = spreadsheet.worksheet("Report")  # Thử lấy sheet "Report"
            print("✅ Sheet 'Report' đã tồn tại. Đang ghi dữ liệu vào sheet.")
        except gspread.exceptions.WorksheetNotFound:
            report_sheet = spreadsheet.add_worksheet(title="Report", rows="100", cols=len(sheet_names_with_data))
            print("✅ Tạo mới sheet 'Report'.")

        # Kiểm tra xem đã có tiêu đề cột chưa
        existing_header = report_sheet.row_values(1)  # Lấy dòng đầu tiên của sheet
        if existing_header != sheet_names_with_data:
            # Tạo header mới theo đúng thứ tự và đầy đủ
            new_header = []
            for col in sheet_names_with_data:
                if col not in existing_header:
                    new_header.append(col)

            # Cập nhật header
            if new_header:
                updated_header = existing_header + new_header
                report_sheet.update('A1', [updated_header])
                print(f"✅ Đã thêm các cột mới: {', '.join(new_header)}")

        # Thêm dữ liệu vào sheet
        max_length = max(len(messages[sheet]) for sheet in sheet_names_with_data)  # Tìm chiều dài dài nhất của các message list

        for i in range(max_length):
            row = []
            for sheet_name in sheet_names_with_data:
                if i < len(messages[sheet_name]):
                    row.append(messages[sheet_name][i])
                else:
                    row.append("")  # Nếu không có message, để trống
            report_sheet.append_row(row)  # Thêm dòng vào sheet

    except Exception as e:
        print(f"❌ Lỗi khi tạo hoặc cập nhật sheet 'Report': {e}")

if __name__ == "__main__":
    # Chạy thực tế
    driver = login()
    open_chat(driver, chat)
    
    current_hour = datetime.datetime.now(pytz.timezone('Asia/Ho_Chi_Minh')).hour

    if current_hour == 14:
        messages = get_filtered_messages(14)
        combined_msgs = combine_messages(messages)  # Gộp message

        print(f"\n✅ Báo cáo lọc được lúc 14h:")
        for sheet_name in sheet_names:
            try:
                print(f"\n{'='*50}")
                print(f"[ {sheet_name} ]")
                print(combined_msgs[sheet_name])
                message = f"[ {sheet_name} ]\n" + combined_msgs[sheet_name]
                send_message(driver, message)
            except:
                continue
        create_or_update_report_sheet(messages)  # Lưu kết quả vào sheet mới

    elif current_hour == 8:
        messages = get_filtered_messages(8)
        combined_msgs = combine_messages(messages)  # Gộp message

        print(f"\n✅ Báo cáo lọc được lúc 8h:")
        for sheet_name in sheet_names:
            try:
                print(f"\n{'='*50}")
                print(f"[ {sheet_name} ]")
                print(combined_msgs[sheet_name])
                message = f"[ {sheet_name} ]\n" + combined_msgs[sheet_name]
                send_message(driver, message)
            except:
                continue
        create_or_update_report_sheet(messages)  # Lưu kết quả vào sheet mới
    
    driver.quit()
    print("✅ Hoàn tất công việc!")
    
