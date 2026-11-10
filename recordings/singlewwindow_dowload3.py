import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(channel="msedge", headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://www.singlewindow.cn/#/")
    page1.get_by_text("中国国际贸易单一窗口").click()
    page1.get_by_text("业务应用", exact=True).click()
    page1.get_by_role("link", name="口岸执法申报").click()
    with page1.expect_popup() as page2_info:
        page1.get_by_role("link", name="加工贸易账册").click()
    page2 = page2_info.value
    page2.locator("a").filter(has_text="综合查询").click()
    page2.get_by_text("保税核注清单查询").click()
    page2.locator("iframe[name=\"iframe04\"]").content_frame.get_by_role("textbox", name="请选择要查询的企业类型").click()
    page2.locator("iframe[name=\"iframe04\"]").content_frame.get_by_role("textbox", name="请选择要查询的企业类型").fill(" ")
    page2.locator("iframe[name=\"iframe04\"]").content_frame.get_by_role("textbox", name="请选择企业类型对应的海关编码；若选择授权企业，仅支持查询当前登录企业申报或录入的清单数据。").click()
    page2.locator("iframe[name=\"iframe04\"]").content_frame.get_by_role("textbox", name="请选择企业类型对应的海关编码；若选择授权企业，仅支持查询当前登录企业申报或录入的清单数据。").fill(" ")
    page2.locator("iframe[name=\"iframe04\"]").content_frame.locator("input[name=\"impExpMarkCdName\"]").click()
    page2.locator("iframe[name=\"iframe04\"]").content_frame.locator("input[name=\"impExpMarkCdName\"]").fill("出口")
    page2.locator("iframe[name=\"iframe04\"]").content_frame.locator("input[name=\"bondInvtNo\"]").click()
    page2.locator("iframe[name=\"iframe04\"]").content_frame.locator("input[name=\"bondInvtNo\"]").fill("2026000000179119283")
    page2.locator("iframe[name=\"iframe04\"]").content_frame.locator("#inputDateStart").click()
    page2.locator("iframe[name=\"iframe04\"]").content_frame.locator("#inputDateStart").click()
    page2.locator("iframe[name=\"iframe04\"]").content_frame.locator("#inputDateStart").fill("20260812")
    page2.locator("iframe[name=\"iframe04\"]").content_frame.locator("#inputDateEnd").click()
    page2.locator("iframe[name=\"iframe04\"]").content_frame.locator("#inputDateEnd").fill("20261110")
    page2.locator("iframe[name=\"iframe04\"]").content_frame.get_by_role("button", name=" 查询").click()
    page2.locator("iframe[name=\"iframe04\"]").content_frame.get_by_text("信息", exact=True).click()
    page2.locator("iframe[name=\"iframe04\"]").content_frame.get_by_text("未查询到相关的清单数据").click()
    page2.locator("iframe[name=\"iframe04\"]").content_frame.get_by_text("确定").click()
    page2.locator("iframe[name=\"iframe04\"]").content_frame.locator("input[name=\"bondInvtNo\"]").click()
    page2.locator("iframe[name=\"iframe04\"]").content_frame.locator("input[name=\"bondInvtNo\"]").fill("QD573026E000091160")
    page2.locator("iframe[name=\"iframe04\"]").content_frame.get_by_role("button", name=" 查询").click()
    page2.locator("iframe[name=\"iframe04\"]").content_frame.get_by_label("", exact=True).first.check()
    with page2.expect_popup() as page3_info:
        page2.locator("iframe[name=\"iframe04\"]").content_frame.get_by_role("button", name="   打印").click()
    page3 = page3_info.value
    page4 = context.new_page()
    page4.goto("edge://downloads-hub/")
    page3.close()
    page4.close()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
