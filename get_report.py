import datetime
from dotenv import load_dotenv
import pytz
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from dateutil import parser
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
# import gc
# gc.disable() # 03/23/2022 added this one to prevent trash collection and avoide crashing the notebooks
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import undetected_chromedriver as uc
from IPython.display import Image, display
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
import time
import json
import os

load_dotenv()

email = os.environ.get('TEAMS_EMAIL')
password = os.environ.get('TEAMS_PASSWORD')
gcp_credentials_json = os.environ.get('GCP_SA_KEY')
chat = "GetReport"
local_tz = pytz.timezone("Asia/Ho_Chi_Minh")
SPREADSHEET_ID = "1_m7s-1-I-SOFfzlWe7CBf5fstFir7qXYAKW4j-8hKYM"
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

#Khởi tạo kết nối
creds_dict = json.loads(gcp_credentials_json)
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, SCOPES)
client = gspread.authorize(creds)
print(f"Spreadsheet ID: {SPREADSHEET_ID}")
spreadsheet = client.open_by_key(SPREADSHEET_ID)

#Lấy danh sách tất cả sheet
sheet_names = [s.title for s in spreadsheet.worksheets()]

# MESSAGE_PATTERN = re.compile(
#     r".*\+\s*([6-9])/.*",
#     re.IGNORECASE | re.DOTALL
# )
MESSAGE_PATTERN = re.compile(r".*\+\s*(\d+)/.*", re.IGNORECASE | re.DOTALL)

def display_screenshot(driver: webdriver.Chrome, file_name: str = 'screenshot.png'):
    """Chụp màn hình và hiển thị"""
    # driver.execute_script("window.scrollBy(0, 250);")
    driver.save_screenshot(file_name)
    time.sleep(3)
    #display(Image(filename=file_name))

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
        time.sleep(2)  # Giữ lại chút delay sau khi gửi
        display_screenshot(driver, "after_sending_message.png")

    except Exception as e:
        print(f"❌ Lỗi khi gửi tin nhắn: {e}")

# def login():
#     import tempfile
#     options = webdriver.ChromeOptions()
#     options.add_argument("--headless")
#     options.add_argument("--no-sandbox")
#     options.add_argument("--disable-dev-shm-usage")
#     options.add_argument("--disable-gpu")
#     options.add_argument("--window-size=1920,1080")
#     temp_dir = tempfile.mkdtemp()
#     options.add_argument(f"--user-data-dir={temp_dir}")

#     driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
#     driver.get("https://teams.live.com/v2/")
#     time.sleep(5)

#     sign_in_btn = WebDriverWait(driver, 20).until(
#         EC.element_to_be_clickable((By.XPATH, '//button[@type="button" and contains(., "Sign in")]'))
#     )
#     sign_in_btn.click()

#     email_input = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "usernameEntry")))
#     email_input.send_keys(email)
#     email_input.send_keys(Keys.RETURN)

#     # Chọn 'Use your password' nếu xuất hiện
#     try:
#         use_pass_btn = WebDriverWait(driver, 10).until(
#             EC.element_to_be_clickable((By.XPATH, '//span[@role="button" and contains(text(), "Use your password")]'))
#         )
#         use_pass_btn.click()
#     except Exception as e:
#         print("Không tìm thấy nút 'Use your password'.")

#     # Tiếp tục nhập mật khẩu như cũ
#     password_input = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "passwordEntry")))
#     password_input.send_keys(password)
#     password_input.send_keys(Keys.RETURN)

#     # driver.save_screenshot("after_email.png")
#     # print("Đã chụp màn hình sau khi nhập email.")

#     try:
#         no_button = WebDriverWait(driver, 15).until(
#             EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="secondaryButton"]'))
#         )
#         no_button.click()
#     except Exception as e:
#         print("Không tìm thấy nút 'No'.")


#     try:
#         button = WebDriverWait(driver, 10).until(
#             EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="primaryButton"]'))
#         )
#         button.click()
#     except:
#         pass
    
#     return driver



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
    #now = datetime.datetime.now(tz)
    now = datetime.datetime.now(tz).replace(tzinfo=None)
    
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
                    else:
                        print(f"Current time là {current_hour}h, không khớp với điều kiện lọc. (8h hoặc 14h)")
                    # else:
                    #     # Lấy tất cả tin nhắn trong vòng 24 tiếng qua ĐỂ TEST
                    #     start_time = now - datetime.timedelta(hours=24)
                    #     end_time = now
                    #     if start_time <= full_datetime <= end_time:
                    #         messages[sheet_name].append(content)

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
        if not sheet_names_with_data:
            print("⚠️ Không có tin nhắn nào thỏa mãn điều kiện lọc. Bỏ qua ghi báo cáo.")
            return
        
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
def get_driver():
    options = uc.ChromeOptions()
    
    # 1. Cấu hình cơ bản cho môi trường Headless (GitHub Actions)
    options.add_argument('--headless=new')
    options.add_argument('--no-sandbox')
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument('--disable-gpu')
    options.add_argument("--window-size=1920,1080")
    
    # 2. [TĂNG TỐC] Chiến lược load trang
    # 'eager' giúp trình duyệt không chờ đợi tải xong ảnh hay script bên thứ 3
    options.page_load_strategy = 'eager'
    
    # 3. [TĂNG TỐC] Chặn tải hình ảnh, CSS và Fonts để tiết kiệm băng thông và RAM
    # prefs = {
    #     "profile.managed_default_content_settings.images": 2,
    #     "profile.managed_default_content_settings.stylesheets": 2,
    #     "profile.managed_default_content_settings.fonts": 2
    # }
    # options.add_experimental_option("prefs", prefs)
    
    # 4. Ép trình duyệt dùng tiếng Anh
    options.add_argument('--lang=en-GB')
    
    # 5. [CHỐNG PHÁT HIỆN] Thêm Proxy dân cư (Khuyến nghị)
    # Trên GitHub Actions, hãy set secrets.PROXY_URL (ví dụ: http://user:pass@ip:port)
    proxy_url = os.getenv("PROXY_URL")
    if proxy_url:
        options.add_argument(f'--proxy-server={proxy_url}')
    
    # Khởi tạo undetected_chromedriver (Không dùng webdriver.Chrome thông thường)
    driver = uc.Chrome(options=options, version_main=146)
    
    # 6. [CHỐNG PHÁT HIỆN] Bơm thêm Stealth Script qua CDP
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            // Ẩn webdriver
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            
            // Fake runtime của Chrome (Bot thường không có cái này)
            window.navigator.chrome = { runtime: {} };
            
            // Bơm thêm plugins giả
            Object.defineProperty(navigator, 'plugins', {
                get: () => [1, 2, 3, 4, 5]
            });
            
            // Ép ngôn ngữ
            Object.defineProperty(navigator, 'languages', {
                get: () => ['en-GB', 'en-US', 'en']
            });
        """
    })
    
    return driver

    
def login():
    driver = get_driver()
    # options = webdriver.ChromeOptions()
    # #options.add_argument("--headless")  
    # options.add_argument("--no-sandbox")
    # options.add_argument("--disable-dev-shm-usage")
    # options.add_argument("--window-size=1920,1080")
    # # Giả lập User-Agent để tránh bị phát hiện là bot
    # options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    
    # temp_dir = tempfile.mkdtemp()
    # options.add_argument(f"--user-data-dir={temp_dir}")

    # driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    # Nếu tài khoản công ty, hãy đổi sang https://teams.microsoft.com/v2/
    driver.get("https://teams.live.com/v2/")
    wait = WebDriverWait(driver, 30)

    try:
        print("⏳ Đang tiến hành đăng nhập...")
        # Bước 1: Click Sign in
        sign_in_btn = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(., "Sign in")]')))
        sign_in_btn.click()
        
        # Bước 2: Nhập Email
        email_input = wait.until(EC.presence_of_element_located((By.ID, "usernameEntry")))
        email_input.send_keys(email)
        email_input.send_keys(Keys.RETURN)
        time.sleep(3)

        # Bước 3: Xử lý nút 'Use your password' nếu xuất hiện
        try:
            use_pass_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Use your password")]'))
            )
            use_pass_btn.click()
        except:
            pass

        # Bước 4: Nhập Password
        pass_input = wait.until(EC.presence_of_element_located((By.ID, "passwordEntry")))
        pass_input.send_keys(password)
        pass_input.send_keys(Keys.RETURN)
        
        # aria-describedby="signIn-title singIn-subtitle"
        # /html/body/div[1]/div/div/div/div[5]/div/div[3]/div/div[1]/div/div/div[1]/div/button[1]
        
        # Bước 5: Vượt qua màn hình 'Stay signed in'
        try:
            no_btn = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="secondaryButton"]'))
            )
            no_btn.click()
        except:
            pass

        print("✅ Đăng nhập thành công!")
        # Đợi giao diện Teams tải xong hoàn toàn
        time.sleep(12) 
        try:
            print("Phát hiện trang bắt Sign in tiếp theo")
            driver.find_element(By.XPATH, '//button[contains(., "Sign in") or contains(@aria-describedby, "signIn-title singIn-subtitle")]').click()
            print("Đã ấn nút Sign in")
            time.sleep(10)
            print("Bắt đầu ấn nút Retry")
            actions = webdriver.ActionChains(driver)
            actions.move_by_offset(500, 500).click().perform()  # Click vào vị trí gần giữa
            actions.send_keys(Keys.TAB).perform()
            time.sleep(1)
            actions.send_keys(Keys.ENTER).perform()
            print("Đã ấn nút Retry")
            time.sleep(20)
            driver.save_screenshot("after_login.png")
        except Exception as e:
            print(f"Không tìm thấy nút đặc thù: {e}")
            driver.save_screenshot("error_login.png")
            print(f"❌ Lỗi đăng nhập: {e}")
            driver.quit()
            return None
        return driver
    except Exception as e:
        driver.save_screenshot("error_login.png")
        print(f"❌ Lỗi đăng nhập: {e}")
        driver.quit()
        return None
    
if __name__ == "__main__":
    # Chạy thực tế
    driver = login()
    time.sleep(5)
    driver.save_screenshot("after_login.png")
    if not driver:
        print("❌ Đăng nhập không thành công!")
    open_chat(driver, chat)

    current_hour = datetime.datetime.now(pytz.timezone('Asia/Ho_Chi_Minh')).hour

    messages = get_filtered_messages(current_hour)
    combined_msgs = combine_messages(messages)  # Gộp message

    print(f"\n✅ Báo cáo lọc được lúc {current_hour}h:")
    for sheet_name in sheet_names:
        try:
            print(f"""Testing
                    Sheet: [ {sheet_name} ]
                    Message: [ {combined_msgs[sheet_name]} ]
                    """)
            # print(f"[ {sheet_name} ]")
            # print(combined_msgs[sheet_name])
            message = f"[ {sheet_name} ]\n" + combined_msgs[sheet_name]
            display_screenshot(driver, f"after_sending.png")
        except:
            continue
    create_or_update_report_sheet(messages)
        
    
    driver.quit()
    print("✅ Hoàn tất công việc!")
    
