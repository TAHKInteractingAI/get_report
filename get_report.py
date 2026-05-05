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
sheet_names = [s.title for s in spreadsheet.worksheets()[:-1]]

# MESSAGE_PATTERN = re.compile(
#     r".*\+\s*([6-9])/.*",
#     re.IGNORECASE | re.DOTALL
# )
MESSAGE_PATTERN = re.compile(r".*(\+\s*\d+/|^\s*\d+\.).*", re.IGNORECASE | re.DOTALL | re.MULTILINE)

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

        display_screenshot(driver, "after_typing_message.png")
        time.sleep(3)
        # Gửi tin nhắn (ENTER)
        message_box.send_keys(Keys.ENTER)
        time.sleep(3)  # Giữ lại chút delay sau khi gửi
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



def open_chat_by_search(driver, chat_name):
    wait = WebDriverWait(driver, 20)
    try:
        # Lấy ô tìm kiếm
        search_xpath = '//input[@placeholder="Search"] | //input[@aria-label="Search"] | //input[@id="ms-searchux-input"]'
        search = wait.until(EC.presence_of_element_located((By.XPATH, search_xpath)))

        # Gõ tên chat
        search.click()
        search.send_keys(Keys.CONTROL + "a")
        search.send_keys(Keys.BACKSPACE)
        search.send_keys(chat_name)
        
        # Chờ kết quả search xổ ra và bấm chọn
        time.sleep(4)
        ActionChains(driver).send_keys(Keys.ARROW_DOWN).pause(1).send_keys(Keys.ENTER).perform()
        
        print(f"Đang chờ load nhóm: {chat_name}...")
        
        # ==========================================
        # 🛡️ CHỐT CHẶN AN TOÀN: Kiểm tra Chat Header
        # ==========================================
        # Đợi tối đa 10s để tiêu đề nhóm chat thay đổi đúng với tên mình cần
        # Selector này lấy tên nhóm ở thanh trên cùng của Teams
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, f"//span[@data-tid='chat-header-title' and contains(text(), '{chat_name[:10]}')] | //hdiv[contains(@class, 'chat-header')]//span[contains(text(), '{chat_name[:10]}')]"))
            )
            print(f"✅ Đã vào đúng nhóm: {chat_name}")
            time.sleep(2) # Chờ tin nhắn load nốt
            return True
            
        except Exception as e:
            print(f"⚠️ Teams load chậm hoặc không tìm thấy nhóm {chat_name}, bỏ qua để tránh cào nhầm data!")
            return False

    except Exception as e:
        print(f"❌ Lỗi khi tìm kiếm nhóm: {chat_name} - {e}")
        return False
def combine_messages(messages_dict):
    """Gộp các message trong cùng sheet thành một chuỗi"""
    combined = {}
    for sheet_name, msg_list in messages_dict.items():
        if msg_list:
            # Gộp các message cách nhau bởi 2 dòng mới
            combined[sheet_name] = "\n\n".join(msg_list)
    return combined

def preprocess_message(content):
    # Bước 1: Chuẩn hóa gạch đầu dòng và khoảng trắng
    content = re.sub(r"-\s+-", "-", content)
    content = re.sub(r"\s*[+]+\s*(\d+/)\s*", r"\n+ \1 ", content)

    # Bước 2: Tách dòng và lọc bỏ + 6/ đến + 10/
    lines = content.splitlines()
    processed_lines = []

    for line in lines:
        line = line.strip()
        if not line: continue

        # XÓA DÒNG + 6 ĐẾN + 10
        if re.match(r"^\+\s*(6|7|8|9|10)\s*/", line):
            continue

        # Thêm ký tự ẩn ZWSP để tránh Teams auto-format sai
        if re.match(r"^(\+|\d+\.|=>|-)", line):
            line = "\u200B" + line
        processed_lines.append(line)

    # Bước 3: THUẬT TOÁN KHỬ TRÙNG LẶP KHỐI (Duplicate trong 1 tin nhắn)
    changed = True
    while changed:
        changed = False
        n = len(processed_lines)
        for L in range(n // 2, 0, -1):
            for i in range(n - 2 * L + 1):
                if processed_lines[i:i+L] == processed_lines[i+L:i+2*L]:
                    processed_lines = processed_lines[:i+L] + processed_lines[i+2*L:]
                    changed = True
                    break
            if changed: break

    content = "\n".join(processed_lines)
    return content.strip()
    
def is_valid_message(content):
    """Kiểm tra message có đúng định dạng"""
    return bool(MESSAGE_PATTERN.match(content))

# Hàm lọc theo thời gian
def get_filtered_messages(current_hour):
    tz = pytz.timezone('Asia/Ho_Chi_Minh')
    now = datetime.datetime.now(tz).replace(tzinfo=None)
    
    messages = {sheet_name: [] for sheet_name in sheet_names}
    # THÊM "Report" VÀ "GetReport" VÀO DANH SÁCH LOẠI TRỪ ĐỂ TRÁNH ĐỌC LẶP
    EXCLUDED_SHEETS = ["Report", "GetReport", "iX000s iSSale TTS Base.XoắnNỆN50k*CấuTrúcVolunt", "iX000s iSSale gbBOSS AH*AU*cOL*YeuCauTop-iUp*KTra"]

    for sheet_name in sheet_names:
        if sheet_name in EXCLUDED_SHEETS: continue
        try:
            sheet = spreadsheet.worksheet(sheet_name)
            data = sheet.get_all_records()

            for row in data:
                try:
                    full_datetime = datetime.datetime.combine(
                        datetime.datetime.strptime(str(row['DATE']), '%Y-%m-%d').date(),
                        parser.parse(str(row['TIME'])).time()
                    )
                    raw_content = row['CONTENT'].strip()

                    # KIỂM TRA HỢP LỆ TRÊN NỘI DUNG GỐC (Trước khi cắt gọt)
                    if not is_valid_message(raw_content): continue
                    
                    # SAU ĐÓ MỚI CẮT GỌT
                    content = preprocess_message(raw_content)
                    if not content: continue # Nếu sau khi cắt mà không còn gì thì bỏ qua

                    # PHÂN LOẠI GIỜ VÀ LOẠI BỎ TIN NHẮN TRÙNG
                    is_in_time = False
                    if current_hour == 8:
                        start = datetime.datetime.combine(now.date() - datetime.timedelta(days=1), datetime.time(13, 0))
                        end = datetime.datetime.combine(now.date(), datetime.time(1, 0))
                        is_in_time = start <= full_datetime < end
                    elif current_hour == 14:
                        start = datetime.datetime.combine(now.date(), datetime.time(1, 0))
                        end = datetime.datetime.combine(now.date(), datetime.time(13, 0))
                        is_in_time = start <= full_datetime < end
                    else: 
                        # Chế độ TEST: lấy 24h qua nếu chạy ngoài giờ 8h và 14h
                        is_in_time = (now - datetime.timedelta(hours=24)) <= full_datetime <= now

                    if is_in_time and content not in messages[sheet_name]:
                        messages[sheet_name].append(content)
                except: continue
        except Exception as e:
            print(f"❌ Lỗi sheet '{sheet_name}': {e}")
    return messages

# Tạo sheet mới để lưu kết quả
def write_to_sheet(sheet_target_name, messages):
    try:
        # Lọc danh sách các sheet có dữ liệu
        sheet_names_with_data = [name for name in sheet_names if messages[name]]
        if not sheet_names_with_data:
            return
        
        # Mở sheet mục tiêu, nếu chưa có thì tạo mới
        try:
            ws = spreadsheet.worksheet(sheet_target_name)
        except gspread.exceptions.WorksheetNotFound:
            ws = spreadsheet.add_worksheet(title=sheet_target_name, rows="1000", cols="20")
            print(f"--- Đã tạo mới sheet: {sheet_target_name} ---")

        # Lấy dữ liệu hiện tại để tránh ghi trùng
        existing_data = ws.get_all_values()
        
        # Cập nhật tiêu đề cột (Header)
        header = sheet_names_with_data
        if not existing_data or existing_data[0] != header:
            ws.update('A1', [header])
            existing_data = [header]

        max_len = max(len(messages[s]) for s in sheet_names_with_data)
        added_count = 0

        for i in range(max_len):
            row = []
            for sheet_name in sheet_names_with_data:
                msg = messages[sheet_name][i] if i < len(messages[sheet_name]) else ""
                row.append(msg)
            
            # Kiểm tra xem dòng này đã có trong sheet chưa trước khi ghi thêm
            if row not in existing_data:
                ws.append_row(row, value_input_option="USER_ENTERED")
                existing_data.append(row)
                added_count += 1
        
        print(f"✅ Hoàn tất ghi vào sheet [{sheet_target_name}]: Thêm {added_count} dòng mới.")
    except Exception as e:
        print(f"❌ Lỗi khi ghi vào sheet {sheet_target_name}: {e}")
        
        
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
        time.sleep(15) 
        
        # SỬA LẠI ĐOẠN NÀY: Kiểm tra nút an toàn, không thấy thì bỏ qua chứ không quit
        try:
            extra_signin = driver.find_elements(By.XPATH, '//button[contains(., "Sign in") or contains(@aria-describedby, "signIn-title singIn-subtitle")]')
            if len(extra_signin) > 0:
                print("Phát hiện trang bắt Sign in tiếp theo...")
                extra_signin[0].click()
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
            else:
                print("👉 Giao diện Teams đã load thẳng, không có popup chặn, tiếp tục công việc!")
                
        except Exception as e:
            print(f"⚠️ Bỏ qua lỗi check màn hình phụ: {e}")
            
        driver.save_screenshot("after_login_success.png")
        return driver  # LUÔN TRẢ VỀ DRIVER VÌ ĐÃ VÀO ĐƯỢC TEAMS

    except Exception as e:
        driver.save_screenshot("error_login.png")
        print(f"❌ Lỗi đăng nhập chính: {e}")
        driver.quit()
        return None
    
if __name__ == "__main__":
    # Chạy thực tế
    for attempt_login in range(5):
        driver = login()
        if driver:
            print("login thành công")
            break
        else:
            print(f"⚠️ Thử đăng nhập lại lần {attempt_login + 1}/5...")
            time.sleep(2)
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
            message = f"[ {sheet_name} ]\n" + combined_msgs[sheet_name]
            send_message(driver, message)
        except:
            continue
            
    # ĐÃ XÓA HÀM CŨ Ở ĐÂY ĐỂ TRÁNH LỖI

    # 1. Ghi dữ liệu sạch vào tab Report
    write_to_sheet("Report", messages)
    
    # 2. Ghi dữ liệu sạch vào tab GetReport
    write_to_sheet("GetReport", messages)
        
    driver.quit()
    print("✅ Hoàn tất toàn bộ công việc!")
