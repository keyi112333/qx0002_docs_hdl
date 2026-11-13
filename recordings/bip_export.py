import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(channel="msedge", headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("http://218.14.150.108:8088/yonbip/resources/uap/rbac/login/main/index.html")
    page.locator("#forwardFrame").content_frame.get_by_role("textbox", name="Alt+K唤起搜索").click()
    page.locator("#forwardFrame").content_frame.get_by_role("textbox", name="Alt+K唤起搜索").press("ControlOrMeta+V")
    page.locator("#forwardFrame").content_frame.get_by_role("textbox", name="Alt+K唤起搜索").fill("外币汇率-全局")
    with page.expect_popup() as page1_info:
        page.locator("#forwardFrame").content_frame.locator("div").filter(has_text=re.compile(r"^外币汇率-全局$")).click()
    page1 = page1_info.value
    page1.frame_locator().get_by_role("textbox", name="请搜索").click()
    page1.frame_locator().get_by_role("textbox", name="请搜索").fill("E")
    page1.frame_locator().get_by_role("textbox", name="请搜索").press("CapsLock")
    page1.frame_locator().get_by_role("textbox", name="请搜索").fill("E0")
    page1.frame_locator().get_by_role("textbox", name="请搜索").press("CapsLock")
    page1.frame_locator().get_by_role("textbox", name="请搜索").press("CapsLock")
    page1.frame_locator().get_by_role("textbox", name="请搜索").press("CapsLock")
    page1.frame_locator().get_by_role("textbox", name="请搜索").press("CapsLock")
    page1.frame_locator().get_by_role("textbox", name="请搜索").fill("E00")
    page1.frame_locator().get_by_role("textbox", name="请搜索").press("CapsLock")
    page1.frame_locator().get_by_role("textbox", name="请搜索").press("CapsLock")
    page1.frame_locator().get_by_role("textbox", name="请搜索").press("CapsLock")
    page1.frame_locator().get_by_role("textbox", name="请搜索").press("CapsLock")
    page1.frame_locator().get_by_role("textbox", name="请搜索").fill("E001")
    page1.frame_locator().get_by_role("textbox", name="请搜索").press("CapsLock")
    page1.frame_locator().get_by_role("textbox", name="请搜索").press("CapsLock")
    page1.frame_locator().get_by_role("textbox", name="请搜索").press("CapsLock")
    page1.locator("#forwardFrame").content_frame.locator("#mainiframe").content_frame.locator("i").first.click()
    page1.frame_locator().locator("a").filter(has_text="E001 基准汇率方案-中间汇率（记账）").click()
    page1.frame_locator().locator("div").filter(has_text=re.compile(r"^123共233条条跳至页$")).first.click()
    page1.frame_locator().locator("#rc_select_0").click()
    page1.frame_locator().locator("#rc_select_0").press("CapsLock")
    page1.frame_locator().locator("#rc_select_0").fill("5")
    page1.frame_locator().locator("#rc_select_0").press("CapsLock")
    page1.frame_locator().locator("#rc_select_0").press("CapsLock")
    page1.frame_locator().locator("#rc_select_0").press("CapsLock")
    page1.frame_locator().locator("#rc_select_0").press("CapsLock")
    page1.frame_locator().locator("#rc_select_0").fill("50")
    page1.frame_locator().locator("#rc_select_0").press("CapsLock")
    page1.frame_locator().locator("#rc_select_0").press("CapsLock")
    page1.frame_locator().locator("#rc_select_0").press("CapsLock")
    page1.frame_locator().locator("#rc_select_0").press("CapsLock")
    page1.frame_locator().locator("#rc_select_0").fill("500")
    page1.frame_locator().locator("#rc_select_0").press("CapsLock")
    page1.frame_locator().locator("#rc_select_0").press("CapsLock")
    page1.frame_locator().locator("#rc_select_0").press("CapsLock")
    page1.frame_locator().locator("#rc_select_0").press("CapsLock")
    page1.frame_locator().locator("#rc_select_0").fill("5000")
    page1.frame_locator().locator("#rc_select_0").press("CapsLock")
    page1.frame_locator().locator("#rc_select_0").press("CapsLock")
    page1.frame_locator().locator("#rc_select_0").press("CapsLock")
    page1.frame_locator().locator("#rc_select_0").press("Enter")
    page1.frame_locator().get_by_role("button").filter(has_text=re.compile(r"^$")).nth(1).click()
    page1.frame_locator().get_by_role("menuitem", name="输出").click()
    page1.frame_locator().get_by_role("radio", name="输出到EXCEL文件").check()
    with page1.expect_download() as download_info:
        page1.locator("#forwardFrame").content_frame.locator("#mainiframe").content_frame.get_by_role("button", name="确定(Y)").click()
    download = download_info.value
    page2 = context.new_page()
    page2.goto("edge://downloads-hub/")

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
