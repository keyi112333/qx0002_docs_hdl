import re
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(channel="msedge", headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://www.singlewindow.cn/#/")
    page.get_by_role("link", name="登录").click()
    page.get_by_role("link", name="卡介质登录").click()
    page.get_by_role("textbox", name="卡口令").click()
    page.get_by_role("textbox", name="卡口令").fill("ushdauhd")   #示例值非实际值，实际读取window普通凭据，找到对应的凭据名读取输入密码（无需账号）
    page.locator("#agreementCardId > .custom-checkbox-wrapper > .custom-checkbox").click()
    page.locator("#checkboxIntel").check()
    page.get_by_role("button", name="登    录").click()
    page.goto("https://www.singlewindow.cn/#/")
    page.get_by_text("业务应用", exact=True).click()
    page.get_by_role("link", name="口岸执法申报").click()
    with page.expect_popup() as page1_info:
        page.get_by_role("link", name="货物申报").click()
    page1 = page1_info.value
    page1.get_by_text("提示", exact=True).click()#可能会出现，出现则点击关闭
    page1.locator("#layui-layer3").get_by_text("关闭").click()
    page1.get_by_text("修改密码：").click()#可能会出现，出现点击已了解，继续访问应用系统
    page1.get_by_text("已了解，继续访问应用系统").click()
    page1.locator("a").filter(has_text="综合查询").first.click()
    page1.get_by_text("报关数据查询").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)