import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(channel="msedge", headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://www.singlewindow.cn/#/")
    page.get_by_role("link", name="登录").click()
    page1.locator("iframe[name=\"iframe01\"]").content_frame.get_by_role("link", name="E20260001895484737").click()
    page1.locator("iframe[name=\"iframe132\"]").content_frame.get_by_text("信息", exact=True).click()
    page1.locator("iframe[name=\"iframe132\"]").content_frame.get_by_text("提示：涉及法定检验检疫要求的出口商品，应先在报关单随附单证栏中填写有效的电子底账数据号，并填写代码B").click()
    page1.locator("iframe[name=\"iframe132\"]").content_frame.get_by_role("button", name="  随附单据").click()
    with page1.expect_download() as download_info:
        page1.locator("iframe[name=\"iframe132\"]").content_frame.get_by_role("cell", name="发票", exact=True).click()
    download = download_info.value
    page2 = context.new_page()
    page2.goto("edge://downloads-hub/")
    page2.close()
    page1.locator("iframe[name=\"iframe132\"]").content_frame.get_by_role("cell", name="装箱单").click()
    with page1.expect_download() as download1_info:
        page1.locator("iframe[name=\"iframe132\"]").content_frame.get_by_role("cell", name="531600000002SW000000004295770947.pdf").click()
    download1 = download1_info.value
    page3 = context.new_page()
    page3.goto("edge://downloads-hub/")

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
