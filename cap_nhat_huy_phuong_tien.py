import logging
import traceback
import sys
import os
import subprocess
from playwright.sync_api import sync_playwright
import pandas as pd
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
    logging.info(f"*** Cập nhật trạng thái phương tiện ***")
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

    data_df = pd.read_excel(excel_path, sheet_name="GIAM", dtype=str)
    data_array = list(data_df.itertuples(index=False, name=None))

    logging.info(f"Đọc thông tin đăng nhập: {email} / {password}")
    logging.info(f"Số lượng phương tiện: {len(data_array)}")

    with sync_playwright() as p:
        logging.info(f"*** Khởi tạo trang ***")
        # show_browser quyết định headless hay không
        browser = p.chromium.launch(headless=not show_browser)
        context = browser.new_context()
        page = context.new_page()

        logging.info("Đang đăng nhập…")
        page.goto("https://qlvh.khaservice.com.vn/login")
        page.locator("//input[@name='email']").fill(email)
        page.locator("//input[@name='password']").fill(password)
        page.locator("//button[@type='submit']").click()
        page.locator("//a[@href='/vehicles']").click()
        page.wait_for_timeout(2000)
        logging.info("=========================================")
        logging.info("*** Bắt đầu cập nhật trạng thái phương tiện ***")

        logging.info(f"=========================================")
        logging.info(f"*** Bắt đầu ***")
        for idx, (phuongtien,) in enumerate(data_array, start=1):
            logging.info(f"Đang xử lý phương tiện mang biển số: {phuongtien}")
            page.locator("//*[@id='input-search-list-style1']").fill(str(phuongtien))
            page.wait_for_timeout(1000)
            try:
                page.wait_for_selector("//*[@data-testid='VisibilityOutlinedIcon']",
                                       state="visible", timeout=3000)
                page.locator("//*[@data-testid='VisibilityOutlinedIcon']").nth(1).click()
            except PlaywrightTimeoutError:
                logging.warning(f"Không thấy phương tiện {phuongtien}, bỏ qua.")
                page.locator("//*[@id='input-search-list-style1']").fill("")
                continue

            page.wait_for_selector("//*[@data-testid='EditOutlinedIcon']").click()
            page.locator("//*[@id='simple-tabpanel-0']/div/div/div[2]/form/div/div[2]/div/div[4]/div[2]").click()
            page.locator("//*[@data-value='Cancelled']").click()
            page.locator("//*[@data-testid='SaveOutlinedIcon']").click()
            page.locator("//*[@data-testid='ArrowBackIosNewIcon']").click()
            # Xóa input tìm kiếm cho vòng lặp tiếp theo
            page.locator("//*[@id='input-search-list-style1']").fill("")
            page.wait_for_timeout(500)

        logging.info("=========================================")
        logging.info("Hoàn tất cập nhật trạng thái phương tiện.")
        browser.close()

if __name__ == "__main__":
    main()




