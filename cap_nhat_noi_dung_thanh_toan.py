import os
import sys
import subprocess
import logging
import pandas as pd
from playwright.sync_api import sync_playwright, Page
import pytest

try:
    from playwright._impl._installer import install

    install("chromium")
except Exception:
    try:
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            check=True
        )
    except Exception as e:
        print("Không thể tải Chromium:", e)
        sys.exit(1)

if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

excel_path = os.path.join(BASE_DIR, "data.xlsx")
project_df = pd.read_excel(excel_path, sheet_name="Project", header=None)
description_val = project_df.iloc[1, 0]
project_list = project_df.iloc[3:, 0].tolist()


# --- Fixtures Pytest ---
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


# --- Test Case Chính ---
def test_cap_nhat_danh_muc(page: Page):
    # 1. Đăng nhập
    page.goto("https://qlvh.khaservice.com.vn/login")
    page.locator("input[name='email']").fill("admin@khaservice.com.vn")
    page.locator("input[name='password']").fill("Admin@123456")
    page.locator("button[type='submit']").click()

    # 2. Vòng lặp cập nhật danh mục
    for idx, project_val in enumerate(project_list, start=1):
        print(f"[{idx}] Project={project_val}")
        logging.error(f"[{idx}] - Project={project_val}")
        page.locator("//*[@id='combo-box-demo']").click()
        page.locator("//*[@id='combo-box-demo']").fill(str(project_val))
        page.locator("//*[@id='combo-box-demo-option-0']").click()
        page.locator("a[href='/configurations/payment']").click()
        page.locator("//a[@href='/configurations/payment']").click()
        page.locator("//*[@data-testid='VisibilityOutlinedIcon']").nth(0).click()
        page.locator("//*[@data-testid='BorderColorIcon']").click()
        page.locator("//textarea[@name='description']").fill("")
        page.locator("//textarea[@name='description']").fill(str(description_val))
        page.locator("//*[@data-testid='SaveOutlinedIcon']").click()
        page.locator("//*[@data-testid='ArrowBackIosNewIcon']").click()

    page.close()