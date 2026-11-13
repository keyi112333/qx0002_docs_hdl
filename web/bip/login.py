#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""BIP 登录与工作台菜单搜索。

从 BIP 销售发票下载脚本中提取，仅保留两段页面操作：

1. 登录：登录页切换环境 → 用户名/密码登录 → 处理提示弹窗；
2. 工作台搜索：Alt+K 搜索并打开菜单页（默认【销售报关单】），返回弹出的业务页面。

筛选、下载、附件解包等逻辑不在本模块内。
"""
import re

from config.credentials import get_credential
from config.config_loader import ConfigError
from log.loggingConfig import logger
from web.browser import BrowserManager

LOGIN_PATH = "/yonbip/resources/uap/rbac/login/main/index.html"
DEFAULT_LOGIN_ENV = "谦信集团正式环境"
MENU_NAME = "销售报关单"
# 工作台顶栏搜索框：不同环境/版本的可访问名可能带空格或额外字符。
SEARCH_BOX_PATTERN = re.compile(r"Alt\+K.*搜索")


def _as_bool(value, default: bool = False) -> bool:
    """配置表里的布尔值可能是 True/False 字符串或“是/否”。"""
    if value is None or str(value).strip() == "":
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in ("1", "true", "yes", "y", "是")


class BipLogin:
    """BIP【登录 + 工作台搜索】页面操作。"""

    def __init__(self, config, browser=None, login_url: str = "",
                 credential_name: str = ""):
        self.config = config
        self.browser = browser
        self._owns_browser = browser is None
        # 登录地址延迟到 open() 解析：未配置BIP时仅创建对象不应报错。
        self._login_url_override = str(login_url or "").strip()
        self.credential_name = str(
            credential_name or config.get("系统配置.BIP凭据目标名") or "").strip()
        self.page = None

    @property
    def login_url(self) -> str:
        return self._login_url_override or self._default_login_url()

    def _default_login_url(self) -> str:
        explicit = str(self.config.get("系统配置.BIP登录地址") or "").strip()
        if explicit:
            return explicit
        base = str(self.config.get("系统配置.BIP地址") or "").strip().rstrip("/")
        if not base:
            raise ConfigError("缺少配置: 系统配置.BIP地址（或系统配置.BIP登录地址）")
        return base + LOGIN_PATH

    def open(self):
        """启动浏览器并登录；同一实例重复调用只置前，不重复登录。"""
        if self.page is not None and not self.page.is_closed():
            try:
                self.page.bring_to_front()
            except Exception:
                pass
            return self.page
        if self.browser is None:
            self.browser = BrowserManager(
                headless=_as_bool(self.config.get("系统配置.无头模式"), False),
                browser_type="chromium",
                channel=str(self.config.get("系统配置.浏览器通道") or "").strip() or None,
            )
            self.browser.start()
        self.page = self.browser.get_page()
        self.page.goto(self.login_url)
        self._login()
        return self.page

    def _login(self) -> None:
        if not self.credential_name:
            raise ConfigError("缺少配置: 系统配置.BIP凭据目标名")
        credential = get_credential(self.credential_name)
        frame = self.page.frame_locator("#forwardFrame")
        self._switch_login_env(frame)
        try:
            frame.get_by_role("textbox", name="用户名").fill(credential.username)
            frame.get_by_role("textbox", name="密码").fill(credential.password)
            frame.get_by_role("button", name="登录").click()
        except Exception as exc:
            raise RuntimeError(f"BIP登录操作失败: {exc}") from exc
        try:
            confirm = frame.get_by_role("button", name="确定").first
            confirm.wait_for(state="visible", timeout=8000)
            confirm.click(timeout=5000)
        except Exception:
            pass  # 无提示信息弹窗时正常进入工作台
        try:
            self._activate_workbench_search()
        except Exception as exc:
            raise RuntimeError("BIP登录后未进入工作台（未找到搜索框）") from exc

    def search_menu(self, name: str = MENU_NAME):
        """工作台 Alt+K 搜索并打开菜单，返回弹出的业务页面。"""
        self._activate_workbench_search().click()
        self.page.wait_for_timeout(400)
        last_error = None
        for attempt in range(2):
            if attempt == 0:
                # 点击后搜索框常被重建/接管，先按键盘输入（焦点在搜索框内）
                self.page.keyboard.type(name, delay=50)
            else:
                try:
                    self._workbench_frame().get_by_role(
                        "textbox", name=SEARCH_BOX_PATTERN).first.fill(name, timeout=8000)
                except Exception as exc:
                    logger.warning("BIP搜索框重新定位后输入失败: %s", exc)
            self.page.wait_for_timeout(1000)
            try:
                with self.page.expect_popup(timeout=60000) as info:
                    self._find_menu_item(name).click(timeout=15000)
                page = info.value
                page.wait_for_load_state()
                logger.info("BIP已打开菜单: %s", name)
                return page
            except Exception as exc:
                last_error = exc
                logger.warning("BIP搜索菜单【%s】第%d次点击失败: %s", name, attempt + 1, exc)
        raise RuntimeError(f"BIP打开【{name}】失败: {last_error}")

    def _switch_login_env(self, frame) -> None:
        """登录页初始环境不固定，必须切换为配置的正式环境（默认谦信集团正式环境）。

        登录页可能有多个 [id$=select] 控件（集团/账套/环境），环境控件的当前值
        通常包含“环境”字样；也可能初始就是别的环境（如天诚溶剂正式环境）。
        """
        env = str(self.config.get("系统配置.BIP登录环境") or DEFAULT_LOGIN_ENV).strip()
        if not env:
            return
        self.page.wait_for_timeout(800)
        selects = frame.locator("[id$='select']")
        try:
            total = selects.count()
        except Exception:
            total = 0
        if total == 0:
            logger.warning("BIP登录页未找到环境下拉控件，按默认环境继续")
            return
        env_selects, other_selects = [], []
        for index in range(total):
            select = selects.nth(index)
            try:
                if "环境" in select.inner_text(timeout=1000):
                    env_selects.append(select)
                else:
                    other_selects.append(select)
            except Exception:
                continue
        for select in env_selects + other_selects:
            try:
                if not select.is_visible():
                    continue
                select.click(timeout=3000)
            except Exception:
                continue
            if not self._click_visible_text(frame, env):
                try:
                    self.page.keyboard.press("Escape")
                except Exception:
                    pass
                continue
            self.page.wait_for_timeout(800)
            try:
                confirmed = env in select.inner_text(timeout=2000)
            except Exception:
                confirmed = False
            if confirmed:
                logger.info("BIP登录环境已切换: %s", env)
            else:
                logger.info("BIP登录环境已选择: %s（未能从控件文本确认）", env)
            return
        logger.warning("BIP登录环境未切换为 %s，请检查登录页环境下拉", env)

    @staticmethod
    def _click_visible_text(frame, text: str, timeout: int = 3000) -> bool:
        """点击当前可见的文本项（环境下拉选项），优先精确匹配。"""
        options = frame.get_by_text(text, exact=True)
        try:
            total = options.count()
        except Exception:
            total = 0
        for index in range(total):
            option = options.nth(index)
            try:
                if option.is_visible():
                    option.click(timeout=timeout)
                    return True
            except Exception:
                continue
        # 选项文本带前后缀时，取最短的可见候选项（最接近选项本身）。
        options = frame.get_by_text(text, exact=False)
        try:
            total = options.count()
        except Exception:
            return False
        candidates = []
        for index in range(total):
            option = options.nth(index)
            try:
                if option.is_visible():
                    candidates.append((len(option.inner_text(timeout=500)), option))
            except Exception:
                continue
        for _, option in sorted(candidates, key=lambda item: item[0]):
            try:
                option.click(timeout=timeout)
                return True
            except Exception:
                continue
        return False

    def _find_menu_item(self, name: str):
        """搜索结果可能在 #forwardFrame，也可能被门户浮层挂到顶层文档。"""
        pattern = re.compile(rf"^{re.escape(name)}$")
        candidate = self._workbench_frame().locator("div").filter(has_text=pattern).first
        try:
            candidate.wait_for(state="visible", timeout=5000)
            return candidate
        except Exception:
            return self.page.locator("div").filter(has_text=pattern).first

    def _workbench_frame(self):
        return self.page.frame_locator("#forwardFrame")

    def _activate_workbench_search(self, timeout: int = 60000):
        """工作台置前并等待顶栏 Alt+K 搜索框可见。"""
        self.page.bring_to_front()
        try:
            self.page.wait_for_load_state("domcontentloaded", timeout=15000)
        except Exception:
            pass
        box = self._workbench_frame().get_by_role(
            "textbox", name=SEARCH_BOX_PATTERN).first
        box.wait_for(state="visible", timeout=timeout)
        return box

    def close(self) -> None:
        if self._owns_browser and self.browser is not None:
            try:
                self.browser.close()
            except Exception as exc:
                logger.debug("关闭BIP浏览器失败: %s", exc)
            self.browser = None
