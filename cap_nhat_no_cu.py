import logging
import traceback
import sys
import os
import subprocess
from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import datetime
import calendar
from dateutil.relativedelta import relativedelta
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

# -----------------------------------------------------
# Đường dẫn file Excel
# === Cấu hình logging ra màn hình ===
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)

excel_path = os.path.join(os.getcwd(), "data.xlsx")
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(os.getcwd(), "ms-playwright")


# -----------------------------------------------------
# Đọc dữ liệu Excel
# -----------------------------------------------------
def ensure_playwright_browsers():
    logging.info(f"=========================================")
    logging.info(f"*** Cập nhật thông tin nợ cũ báo phí ***")
    logging.info(f"=========================================")
    browsers_path = os.environ["PLAYWRIGHT_BROWSERS_PATH"]
    chromium_exists = False

    if os.path.exists(browsers_path):
        for d in os.listdir(browsers_path):
            if d.startswith("chromium-"):
                chromium_exists = True
                break

    if chromium_exists:
        logging.info("Đã tìm thấy Chromium tại %s", browsers_path)
    else:
        logging.info("Không tìm thấy Chromium. Đang tải…")
        try:
            subprocess.run(
                [sys.executable, "-m", "playwright", "install", "chromium"],
                check=True
            )
            logging.info("Đã tải xong Chromium.")
        except Exception as e:
            logging.error("Không thể tải Chromium: %s", e)


# -----------------------------------------------------
# Đảm bảo Chromium Playwright đã cài
# -----------------------------------------------------

def main():
    try:
        ensure_playwright_browsers()
        print("Bạn có muốn xem quá trình thực hiện không? Y / N")
        print("Nếu Y thì tắt bớt những chương trình khác để máy chạy mượt hơn!")
        choice = input("Nhập (Y/N). Sau đó nhấn Enter: ").strip().lower()
        show_browser = True if choice == 'y' else False

        run_test(show_browser=show_browser)
        logging.info("Đã hoàn thành.")
    except Exception as e:
        logging.error("Đã xảy ra lỗi:")
        traceback.print_exc()
    finally:
        input("\nNhấn Enter để thoát...")


def run_test(show_browser=False):
    """
    show_browser = True  => hiện giao diện trình duyệt
    show_browser = False => chạy headless
    """
    login_df = pd.read_excel(excel_path, sheet_name="Login")
    email = login_df.iloc[0, 0]
    password = login_df.iloc[0, 1]

    data_df = pd.read_excel(excel_path, sheet_name="NoCu", dtype=str)
    data_array = list(data_df.itertuples(index=False, name=None))

    logging.info(f"Đọc thông tin đăng nhập: {email} / {password}")
    logging.info(f"Số lượng căn hộ cập nhật: {len(data_array)}")

    with sync_playwright() as p:
        logging.info(f"*** Khởi tạo trang ***")
        # show_browser quyết định headless hay không
        browser = p.chromium.launch(
            headless=not show_browser,
            args=["--start-maximized"] if show_browser else []
        )
        # bỏ viewport mặc định => trang sẽ theo kích thước cửa sổ thật (full màn hình)
        context = browser.new_context(no_viewport=True)
        page = context.new_page()

        logging.info("*** Đang đăng nhập ***")
        page.goto("https://qlvh.khaservice.com.vn/login")
        page.locator("//input[@name='email']").fill(email)
        page.locator("//input[@name='password']").fill(password)
        page.locator("//button[@type='submit']").click()
        page.locator("//a[@href='/fee-reports']").click()
        page.wait_for_timeout(2000)
        logging.info("=========================================")
        logging.info("*** Bắt đầu cập nhật nợ cũ căn hộ ***")

        logging.info(f"=========================================")
        for idx, (canho, thang, sotien) in enumerate(data_array, start=1):
            logging.info(f"[{idx}/{len(data_array)}] Căn hộ: {canho} - Tháng: {thang}")
            canho_str = str(canho)
            if len(canho_str) == 3:
                canho_str = "0" + canho_str
            page.locator("//*[@id='root']/div[2]/main/div/div/div[1]/div/span/div/div[2]/div/button[2]").click()
            page.locator("//html/body/div[2]/div[3]/div/div[1]/div/button").click()
            page.locator("//html/body/div[2]/div[3]/div/div[2]/div/div[2]/div/input").fill(canho_str)
            page.locator("//*[@data-option-index='0']").click()
            # Chọn tháng
            page.locator("//*[@placeholder='MM/YYYY']").fill(str(thang))
            page.keyboard.press("Escape")  # ẩn lịch
            # Điền căn hộ
            # page.locator("//*[@id='input-search-list-style1']").fill(str(canho))
            # page.wait_for_timeout(1000)  # chờ kết quả

            try:
                page.wait_for_selector("//*[@data-testid='VisibilityOutlinedIcon']",
                                       state="visible", timeout=3000)
                page.locator("//*[@data-testid='VisibilityOutlinedIcon']").nth(1).click()
            except PlaywrightTimeoutError:
                logging.warning(f"Không thấy căn hộ: {canho}, bỏ qua.")
                page.locator("//*[@id='input-search-list-style1']").fill("")
                continue

            # Các thao tác khác
            page.locator("//*[@id='simple-tabpanel-0']/div/div/div/form/div[2]/div[13]/div/div/button").click()

            page.wait_for_timeout(1000)
            dt = datetime.strptime(thang, "%m/%Y")
            thang_truoc = dt - relativedelta(months=1)
            thang_truoc_str = thang_truoc.strftime("%m/%Y")
            page.locator("//html/body/div[2]/div[3]/div/div[1]/div/div[1]//input").fill("Nợ cũ /The debt of" + " " + thang_truoc_str)
            page.locator("//html/body/div[2]/div[3]/div/div[1]/div/div[2]/div/div").click()
            page.locator("//*[@data-value='Olddebit']").click()
            page.locator("//html/body/div[2]/div[3]/div/div[1]/div/div[3]//input").fill("1")
            page.locator("//html/body/div[2]/div[3]/div/div[1]/div/div[4]//input").fill("Tháng/ Month")
            page.locator("//html/body/div[2]/div[3]/div/div[1]/div/div[5]//input").fill(str(sotien))
            page.locator("//html/body/div[2]/div[3]/div/div[1]/div/div[6]//input").fill("VND")
            page.locator("//html/body/div[2]/div[3]/div/div[1]/div/div[9]/div").nth(1).click()
            page.locator("//*[@data-value='0']").click()
            page.locator("//html/body/div[2]/div[3]/div/div[1]/div/div[10]//input").fill(thang_truoc_str)
            ngay_dau = thang_truoc.replace(day=1)
            ngay_cuoi = thang_truoc.replace(
                day=calendar.monthrange(thang_truoc.year, thang_truoc.month)[1]
            )
            ngay_dau_str = ngay_dau.strftime("%d/%m/%Y")
            ngay_cuoi_str = ngay_cuoi.strftime("%d/%m/%Y")
            page.locator("//html/body/div[2]/div[3]/div/div[1]/div/div[11]//input").fill(ngay_dau_str)
            page.locator("//html/body/div[2]/div[3]/div/div[1]/div/div[12]//input").fill(ngay_cuoi_str)
            page.locator("//button[@type='submit']").click()
            page.locator("//*[@data-testid='SaveOutlinedIcon']").click()
            logging.info(f"Đã xử lý căn hộ: {canho} - Tháng: {thang}")
            page.locator("//*[@data-testid='ArrowBackIosNewIcon']").click()
            page.wait_for_timeout(500)
        logging.info(f"=========================================")
        page.close()

if __name__ == "__main__":
    main()




