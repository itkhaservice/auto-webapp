import logging
import traceback
import sys
import os
import subprocess

from openpyxl.styles.fills import fills
from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import datetime
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
    logging.info(f"*** Cập nhật phieu thu tien mat ***")
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

    data_df = pd.read_excel(excel_path, sheet_name="TIENMAT", dtype=str)
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
        page.locator("//a[@href='/receipt-vouchers']").click()
        page.wait_for_timeout(2000)
        logging.info("=========================================")
        logging.info("*** Bắt đầu cập nhật Đã thanh toán trước căn hộ ***")

        logging.info(f"=========================================")

        for idx, (canho, block, ngay, ten) in enumerate(data_array, start=1):
            logging.info(f"[{idx}/{len(data_array)}] Căn hộ: {canho}")

            page.locator("//*[@id='root']/div[2]/main/div/div/div[1]/div/span/div/div[2]/div/div[1]/button").click()

            page.locator("//*[@id='root']/div[2]/main/div/div/div/form/div/div[2]/div/div[6]/div/div[2]/div").click()

            if block == "BLOCK B":
                page.locator("//*[@data-option-index='0']").click()
            elif block == "BLOCK A":
                page.locator("//*[@data-option-index='1']").click()
            else:
                logging.warning(f"Block không hợp lệ: {block}")
            logging.info(f"Ngày: {str(ngay)}")

            ngay_da_dinh_dang = ngay.strftime("%d/%m/%Y")
            page.locator("//*[@placeholder='DD/MM/YYYY']").fill(ngay_da_dinh_dang)

            canho_str = str(canho)
            if len(canho_str) == 3:
                canho_str = "0" + canho_str
            page.locator("//*[@id='root']/div[2]/main/div/div/div/form/div/div[2]/div/div[7]/div/div[2]/div/input").fill(str(canho_str))
            page.locator("//*[@id='combo-box-demo-listbox']").click()

            page.locator("//*[@id='root']/div[2]/main/div/div/div/form/div/div[2]/div/div[9]/div[2]/div/input").fill(str(ten))

            selector_input = 'input[type="checkbox"][data-indeterminate="false"]'
            page.check(selector_input)

            page.locator("//*[@id='root']/div[2]/main/div/div/div/form/div/div[1]/div/h5/div/div[2]/div/button").click()
            page.wait_for_timeout(500)
        logging.info(f"=========================================")
        page.close()

if __name__ == "__main__":
    main()




