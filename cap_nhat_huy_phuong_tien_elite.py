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
excel_path = os.path.join(os.getcwd(), "data.xlsx")
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(os.getcwd(), "ms-playwright")


def ensure_playwright_browsers():
    logging.info(f"=========================================")
    logging.info(f"*** Cập nhật trạng thái thanh toán ***")
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
def ensure_playwright_browsers():
    logging.info("Kiểm tra trình duyệt Chromium…")
    try:
        from playwright._impl._installer import install
        install("chromium")
        logging.info("Chromium đã có sẵn.")
    except Exception:
        logging.info("Không tìm thấy Chromium. Đang tải…")
        try:
            subprocess.run(
                [sys.executable, "-m", "playwright", "install", "chromium"],
                check=True
            )
            logging.info("Đã tải xong Chromium.")
        except Exception as e:
            logging.error("Không thể tải Chromium: %s", e)
            sys.exit(1)

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
        page.locator("//*[@id='root']/div[2]/main/div/div/div[1]/div/span/div/div[2]/div/button[3]").click()
        page.locator('//*[@id="outlined-basic1"]').nth(0).click()
        page.locator('//*[@data-value="Using"]').wait_for(state="visible")
        page.locator('//*[@data-value="Using"]').click()
        page.keyboard.press("Escape")
        for idx, (phuongtien,) in enumerate(data_array, start=1):
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
            page.locator("//*[@data-testid='VisibilityOutlinedIcon']").click()
            canho_text = page.locator(
                '//*[@id="simple-tabpanel-0"]/div/div/div[2]/form/div/div[2]/div/div[21]/p/p'
            ).inner_text()
            loaixe_text = page.locator(
                '//*[@id="simple-tabpanel-0"]/div/div/div[2]/form/div/div[2]/div/div[22]/p/p'
            ).inner_text()
            logging.info(
                f"Căn hộ: {canho_text} - Đã hủy phương tiện mang biển số: {phuongtien} - Loại xe: {loaixe_text}")
            if loaixe_text.strip() == "Xe máy 1/ Motorbike 1":
                page.locator("//*[@data-testid='ArrowBackIosNewIcon']").click()
                page.locator("//*[@id='input-search-list-style1']").fill(str(canho_text))
                page.locator("//*[@id='root']/div[2]/main/div/div/div[1]/div/span/div/div[2]/div/button[3]").click()
                page.locator('//*[@id="outlined-basic1"]').nth(2).click()
                page.locator('//*[@id="menu-"]/div[3]/ul/li[1]').click()
                page.keyboard.press("Escape")
                try:
                    page.wait_for_selector("//*[@data-testid='VisibilityOutlinedIcon']",
                                           state="visible", timeout=3000)
                    page.locator("//*[@data-testid='VisibilityOutlinedIcon']").nth(1).click()
                    page.locator("//*[@data-testid='EditOutlinedIcon']").click()
                    page.locator("//*[@id='outlined-basic1']").nth(2).click()
                    page.locator("//*[@id='menu-']/div[3]/ul/li[3]").click()
                    page.locator("//*[@data-testid='SaveOutlinedIcon']").click()
                    page.locator("//*[@data-testid='VisibilityOutlinedIcon']").click()
                    loaixe_moi_text = page.locator(
                        '//*[@id="simple-tabpanel-0"]/div/div/div[2]/form/div/div[2]/div/div[3]/p/p'
                    ).inner_text()
                    logging.info(
                        f"Căn hộ: {canho_text} - đã thay đổi phương tiện mang biển số: {loaixe_moi_text} - Thành loại xe: Xe máy 1/ Motorbike 1")

                    page.locator("//*[@data-testid='ArrowBackIosNewIcon']").click()
                except PlaywrightTimeoutError:
                    logging.info(
                        f"Căn hộ: {canho_text} - đã xử lý phương tiện mang biển số: {phuongtien} - Loại xe: {loaixe_text}")
                    continue

            else:
                logging.info(
                    f"Căn hộ: {canho_text} - đã xử lý phương tiện mang biển số: {phuongtien} - Loại xe: {loaixe_text}")
                page.locator("//*[@data-testid='ArrowBackIosNewIcon']").click()
                # Xóa input tìm kiếm cho vòng lặp tiếp theo
                page.locator("//*[@id='input-search-list-style1']").fill("")
                page.wait_for_timeout(500)


        logging.info("=========================================")
        logging.info("Hoàn tất cập nhật trạng thái phương tiện.")
        browser.close()

if __name__ == "__main__":
    main()




