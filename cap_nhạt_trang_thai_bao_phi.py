import os
import sys
import subprocess
import logging
import pandas as pd
from playwright.sync_api import sync_playwright, Page
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
import traceback

# --- Đảm bảo Chromium Playwright được tải ---
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

def ensure_playwright_browsers():
    logging.info(f"=========================================")
    logging.info(f"*** Cập nhật trạng thái báo phí ***")
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

def run_test(show_browser=False):
    login_df = pd.read_excel(excel_path, sheet_name="Login")
    email = login_df.iloc[0, 0]
    password = login_df.iloc[0, 1]

    data_df = pd.read_excel(excel_path, sheet_name="Data", dtype=str)
    data_array = list(data_df.itertuples(index=False, name=None))

    logging.info(f"Đọc thông tin đăng nhập: {email} / {password}")
    logging.info(f"Số lượng căn hộ: {len(data_array)}")

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
        # Đăng nhập
        page.goto("https://qlvh.khaservice.com.vn/login")
        page.locator("//input[@name='email']").fill(email)
        page.locator("//input[@name='password']").fill(password)
        page.locator("//button[@type='submit']").click()
        page.locator("//a[@href='/fee-reports']").click()


        for idx, (canho, thang) in enumerate(data_array, start=1):
            logging.info(f"[{idx}/{len(data_array)}] Căn hộ: {canho} - Tháng: {thang}")
            # Chọn tháng
            page.locator("//*[@id='root']/div[2]/main/div/div/div[1]/div/span/div/div[2]/div/button[2]").click()
            page.locator("//*[@placeholder='MM/YYYY']").fill(str(thang))
            page.keyboard.press("Escape")  # ẩn lịch

            page.locator("//*[@id='input-search-list-style1']").fill(str(canho))
            page.wait_for_timeout(1000)  # chờ kết quả

            try:
                page.wait_for_selector("//*[@data-testid='VisibilityOutlinedIcon']",
                                       state="visible", timeout=3000)
                page.locator("//*[@data-testid='VisibilityOutlinedIcon']").nth(1).click()
            except PlaywrightTimeoutError:
                logging.warning(f"Không thấy căn hộ: {canho}, bỏ qua.")
                page.locator("//*[@id='input-search-list-style1']").fill("")
                continue

            # Các thao tác khác
            page.locator("//*[@id='simple-tabpanel-0']/div/div/div/form/div[2]/div[3]/div[2]/div").click()
            page.locator("//*[@data-value='1']").click()

            page.locator("//*[@data-testid='SaveOutlinedIcon']").click()
            page.locator("//*[@data-testid='ArrowBackIosNewIcon']").click()

            # Xóa input tìm kiếm cho vòng lặp tiếp theo
            page.locator("//*[@id='input-search-list-style1']").fill("")
            page.wait_for_timeout(500)

        page.close()

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

if __name__ == "__main__":
    main()