import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(channel="msedge", headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://www.singlewindow.cn/#/")
    page1.locator("iframe[name=\"iframe01\"]").content_frame.get_by_text("一键打印 打印报关单 打印通知书 打印电子委托协议 批量申报 列表导出 详情导出 导入 删除 状态订阅 复制 修改申请 撤销申请 一键预检 重传补传 未完整申报").check()#全选
    """<th class="bs-checkbox " style="width: 36px; " data-field="state" tabindex="0"><div class="th-inner "><input name="btSelectAll" type="checkbox"></div><div class="fht-cell" style="width: 36px;"></div></th>"""
    with page1.expect_download() as download_info:
        page1.locator("iframe[name=\"iframe01\"]").content_frame.get_by_role("button", name="  列表导出").click()
    download = download_info.value
    page2 = context.new_page()
    page2.goto("edge://downloads-hub/")
    with page1.expect_download() as download5_info:
        with page1.expect_popup() as page3_info:
            page1.locator("iframe[name=\"iframe01\"]").content_frame.get_by_role("button", name="  详情导出").click()
        page3 = page3_info.value
    download5 = download5_info.value
    page1.locator("iframe[name=\"iframe01\"]").content_frame.get_by_role("row", name="E20260001897868651").get_by_role("checkbox").check()#name后面是统一编号
    page1.locator("iframe[name=\"iframe01\"]").content_frame.get_by_role("button", name="  打印报关单").click()
    with page1.expect_popup() as page10_info:
        page1.locator("iframe[name=\"iframe01\"]").content_frame.get_by_text("打印预览").click()
    page10 = page10_info.value
    page11 = context.new_page()
    page11.goto("edge://downloads-hub/")
    page10.close()
    page11.close()
    page1.locator("iframe[name=\"iframe01\"]").content_frame.get_by_role("button", name="  打印通知书").click()
    with page1.expect_popup() as page12_info:
        page1.locator("iframe[name=\"iframe01\"]").content_frame.get_by_text("打印预览").click()
    page12 = page12_info.value
    page13 = context.new_page()
    page13.goto("edge://downloads-hub/")
    page12.close()
    page13.close()
    page1.locator("iframe[name=\"iframe01\"]").content_frame.get_by_role("button", name="  打印电子委托协议").click()
    with page1.expect_popup() as page14_info:
        page1.locator("iframe[name=\"iframe01\"]").content_frame.get_by_text("打印预览").click()
    page14 = page14_info.value
    page15 = context.new_page()
    page15.goto("edge://downloads-hub/")
    page14.close()
    page15.close()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
