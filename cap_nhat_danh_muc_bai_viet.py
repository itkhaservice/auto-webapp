import os
import sys
import subprocess
import logging
import pandas as pd
from playwright.sync_api import sync_playwright, Page
import pytest

# --- Đảm bảo Chromium Playwright được tải ---
try:
    from playwright._impl._installer import install
    install("chromium")  # tải Chromium nếu chưa có
except Exception:
    try:
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            check=True
        )
    except Exception as e:
        print("Không thể tải Chromium:", e)
        sys.exit(1)

# --- Đường dẫn file Excel ---
if getattr(sys, 'frozen', False):  # Khi chạy exe
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

excel_path = os.path.join(BASE_DIR, "data.xlsx")

# --- Đọc Excel ---
login_df = pd.read_excel(excel_path, sheet_name="Login")
email = login_df.loc[0, "email"]
password = login_df.loc[0, "password"]

data_df = pd.read_excel(excel_path, sheet_name="Data", dtype=str)
data_array = list(data_df.itertuples(index=False, name=None))


@pytest.fixture(scope="session")
def browser():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=["--start-maximized"])
        yield browser
        browser.close()

@pytest.fixture
def page(browser):
    context = browser.new_context(no_viewport=True)
    page = context.new_page()
    yield page
    context.close()

def test_cap_nhat_danh_muc(page: Page):
    # 1. Đăng nhập
    page.goto("https://qlvh.khaservice.com.vn/login")
    page.locator("//input[@name='email']").fill(email)
    page.locator("//input[@name='password']").fill(password)
    page.locator("//button[@type='submit']").click()

    page.locator("//p[normalize-space(text())='Bài viết']").click()
    # if p_locator.count() == 0:
    #     p_locator = page.locator("//p[normalize-space(text())='Post']")
    #     if p_locator.count() == 0:
    #         raise Exception("Không tìm thấy thẻ <p> chứa 'Bài viết' hoặc 'Post'")
    # p_locator.first.click()

    page.locator("//a[@href='/posts/post-categories']").click()

    data = [
        # Nội quy NCC (3)
        ("Nội quy về sửa chữa, cải tạo căn hộ", "Nội quy về sửa chữa, cải tạo căn hộ", 3),
        ("Nội quy về đỗ xe và phương tiện giao thông", "Nội quy về đỗ xe và phương tiện giao thông", 3),
        ("Nội quy về vật nuôi trong chung cư", "Nội quy về vật nuôi trong chung cư", 3),
        ("Nội quy sử dụng thang máy và khu vực chung", "Nội quy sử dụng thang máy và khu vực chung", 3),
        ("Nội quy về vệ sinh và môi trường", "Nội quy về vệ sinh và môi trường", 3),
        ("Nội quy sinh hoạt và an ninh trật tự", "Nội quy sinh hoạt và an ninh trật tự", 3),
        ("Nội quy chung cư về quyền và nghĩa vụ cư dân", "Nội quy chung cư về quyền và nghĩa vụ cư dân", 3),

        # Quảng cáo (2)
        ("Quảng cáo dịch vụ tiện ích", "Quảng cáo dịch vụ tiện ích", 2),
        ("Quảng cáo sự kiện & khuyến mãi", "Quảng cáo sự kiện & khuyến mãi", 2),

        # Tin tức (1)
        ("Tin tức bảo trì & sửa chữa", "Tin tức bảo trì & sửa chữa", 0),
        ("Tin tức tiện ích & dịch vụ", "Tin tức tiện ích & dịch vụ", 0),
        ("Tin tức cộng đồng cư dân", "Tin tức cộng đồng cư dân", 0),
        ("Tin tức pháp lý & chính sách", "Tin tức pháp lý & chính sách", 0),
        ("Tin tức nội bộ chung cư", "Tin tức nội bộ chung cư", 0),

        # Thông báo (0)
        ("Thông báo nghỉ lễ", "Thông báo nghỉ lễ", 1),
        ("Thông báo bảo trì & sửa chữa", "Thông báo bảo trì & sửa chữa", 1),
        ("Thông báo nội quy & quy định", "Thông báo nội quy & quy định", 1),
        ("Thông báo vệ sinh & môi trường", "Thông báo vệ sinh & môi trường", 1),
        ("Thông báo an ninh & an toàn", "Thông báo an ninh & an toàn", 1),
        ("Thông báo sự kiện & hoạt động", "Thông báo sự kiện & hoạt động", 1),
        ("Thông báo phí & thanh toán", "Thông báo phí & thanh toán", 1),
    ]

    for name, desc, cat in data:
        page.locator("//*[@data-testid='AddOutlinedIcon']").click()
        page.locator("//*[@name='name']").fill(name)
        page.locator("//textarea[@name='description']").fill(desc)
        page.locator('//*[@id="root"]/div[2]/main/div/div/div[2]/div/div[2]/div/div[3]/div[2]/div').click()
        page.locator("//li[@data-value='0']").click()
        page.locator('//*[@id="root"]/div[2]/main/div/div/div[2]/div/div[2]/div/div[2]/div[2]/div').click()
        page.locator(f"//li[@data-value='{cat}']").click()
        page.locator("//*[@data-testid='AddOutlinedIcon']").click()

    page.close()
