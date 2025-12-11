import logging
import pytest
from playwright.sync_api import Page

def test_cap_nhat_loai_can_ho(page: Page) -> None:
    page.goto("https://qlvh.khaservice.com.vn/login")
    page.locator("//input[@name='email']").fill("bql.topazhome2blockb@khaservice.com.vn")
    page.locator("//input[@name='password']").fill("Bqltopazhome2@")
    page.locator("//button[@type='submit']").click()
    page.locator("//a[@href='/apartments']").click()

    # Danh sách các căn hộ cần lặp qua
    canho_raw_list = ["B3.0302","B2.0901","B2.0702","B2.0608","B1.1206"]

    # Đặt vòng lặp for ở đây
    for canho in canho_raw_list:
        logging.info(f"Đang xử lý căn hộ: {canho}") # Thêm log để dễ theo dõi

        # Điền giá trị căn hộ hiện tại vào ô tìm kiếm
        # Chú ý: thay vì .fill("") là .fill(canho)
        page.locator("//*[@id='input-search-list-style1']").fill(canho)

        # Đợi một chút để kết quả tìm kiếm hiển thị (tùy chọn, nhưng thường hữu ích)
        page.wait_for_timeout(1000) # Đợi 1 giây, điều chỉnh nếu cần

        # Các thao tác khác trong vòng lặp
        page.locator("//*[@data-testid='VisibilityOutlinedIcon']").nth(1).click()
        page.locator("//*[@data-testid='EditOutlinedIcon']").click()
        page.locator("//*[@aria-haspopup='listbox']").click()
        page.locator("//*[@data-value='Apartment']").click()
        page.locator("//*[@data-testid='SaveOutlinedIcon']").click()
        page.locator("//*[@data-testid='ArrowBackIosNewIcon']").click()

        # Xóa nội dung ô tìm kiếm sau mỗi lần lặp để chuẩn bị cho lần tìm kiếm tiếp theo
        # Nếu trang không tự động làm mới, bạn có thể cần điều này
        page.locator("//*[@id='input-search-list-style1']").fill("")
        page.wait_for_timeout(500) # Đợi một chút sau khi xóa

    page.close()


def test_cap_nhat_dinh_muc_nhan_khau(page: Page) -> None:
    page.goto("https://qlvh.khaservice.com.vn/login")
    page.locator("//input[@name='email']").fill("bql.topazhome2blockb@khaservice.com.vn")
    page.locator("//input[@name='password']").fill("Bqltopazhome2@")
    page.locator("//button[@type='submit']").click()
    page.locator("//a[@href='/apartments']").click()

    # Dữ liệu căn hộ
    canho_raw_list = [
            {
                "apt": "B1.0709",
                "num": 1
            }
    ]

    for canho in canho_raw_list:
        apt_value = canho["apt"]
        num_value = str(canho["num"])
        logging.info(f"Đang xử lý căn hộ: {apt_value}") # Thêm log để dễ theo dõi
        page.locator("//*[@id='input-search-list-style1']").fill(apt_value)
        page.wait_for_timeout(1000)
        page.locator("//*[@data-testid='VisibilityOutlinedIcon']").nth(1).click()
        page.locator("//*[@data-testid='EditOutlinedIcon']").click()
        page.locator("//*[@name='waterNorm']").fill("")
        page.locator("//*[@name='waterNorm']").fill(num_value)
        page.locator("//*[@data-testid='SaveOutlinedIcon']").click()
        page.locator("//*[@data-testid='ArrowBackIosNewIcon']").click()
        page.locator("//*[@id='input-search-list-style1']").fill("")
        page.wait_for_timeout(500)
    page.close()