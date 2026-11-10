#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
单一窗口-综合查询-报关数据查询
"""
import re
import time
from datetime import date

from log.loggingConfig import logger
from web.base_page import BasePage

IFRAME_NAME = "iframe01"
FORM_ANCHOR = "#etpsCategory1Name"

# 赋值脚本：同时触发 input/change，兼容 jQuery/laydate 控件
JS_SET_VALUE = """
(el, value) => {
    el.value = value;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
}
"""


class DeclarationQuery(BasePage):
    """报关数据查询条件填写与查询

    查询条件由隐藏域 + 文本显示域组成（如 etpsCategory1/etpsCategory1Name），
    文本显示域由 autocompleter 渲染，直接点击文本不可靠，因此按隐藏域 id 赋值。
    页面可能存在多个同名 iframe，按“可见的表单字段”定位实际使用的 frame。
    """

    # (隐藏域id, 文本域id, 隐藏域值, 文本显示值)
    ETPS_CATEGORY = ("etpsCategory1", "etpsCategory1Name", "C", "C-报关收发货人")
    IE_FLAG = ("cusIEFlag1", "cusIEFlag1Name", "E", "E-出口")
    TABLE_FLAG = ("tableFlag", "tableFlagName", "1", "1-是")

    def search(self, start_date: date, end_date: date, submit: bool = True) -> bool:
        """填写查询条件并点击查询

        Args:
            start_date: 最近操作时间起
            end_date: 最近操作时间止（单次查询不超过7天）
            submit: 是否点击查询按钮，False 时只填写条件便于人工确认
        """
        frame = self._find_form_frame()
        logger.info(f"报关数据查询表单: {frame.url}")

        self._ensure_autocompleter(frame, *self.ETPS_CATEGORY)
        self._ensure_autocompleter(frame, *self.IE_FLAG)
        self._ensure_autocompleter(frame, *self.TABLE_FLAG)

        start_text = start_date.strftime("%Y-%m-%d")
        end_text = end_date.strftime("%Y-%m-%d")
        self._set_value(frame, "#updateTime1", start_text)
        self._set_value(frame, "#updateTimeEnd1", end_text)

        self._check_value(frame, "#etpsCategory1", self.ETPS_CATEGORY[2], "企业类别编码")
        self._check_value(frame, "#etpsCategory1Name", self.ETPS_CATEGORY[3], "企业类别")
        self._check_value(frame, "#cusIEFlag1", self.IE_FLAG[2], "进出口标志编码")
        self._check_value(frame, "#cusIEFlag1Name", self.IE_FLAG[3], "进出口标志")
        self._check_value(frame, "#tableFlag", self.TABLE_FLAG[2], "是否结关编码")
        self._check_value(frame, "#tableFlagName", self.TABLE_FLAG[3], "是否结关")
        self._check_value(frame, "#updateTime1", start_text, "最近操作时间起")
        self._check_value(frame, "#updateTimeEnd1", end_text, "最近操作时间止")

        logger.info(
            f"查询条件已填写: 企业类别={self.ETPS_CATEGORY[3]}，"
            f"进出口标志={self.IE_FLAG[3]}，是否结关={self.TABLE_FLAG[3]}，"
            f"最近操作时间 {start_text} ~ {end_text}"
        )
        if not submit:
            logger.info("已填写查询条件（未提交）")
            return True

        frame.get_by_role("button", name=re.compile("查询")).first.click()
        logger.info("已提交报关数据查询")
        return True

    def _find_form_frame(self, timeout: int = 30000):
        """在所有 frame 中查找可见的报关数据查询表单

        报关数据查询页是点击菜单后异步加载的，需要轮询等待 iframe 和表单渲染完成。
        """
        deadline = time.time() + timeout / 1000
        while time.time() < deadline:
            for frame in self.page.frames:
                try:
                    anchor = frame.locator(FORM_ANCHOR)
                    if anchor.count() > 0 and anchor.first.is_visible():
                        return frame
                except Exception:
                    continue
            self.page.wait_for_timeout(500)

        candidates = []
        visible_inputs = []
        for frame in self.page.frames:
            try:
                count = frame.locator(FORM_ANCHOR).count()
                if count:
                    candidates.append(
                        f"{frame.name}|{frame.url}|count={count}|"
                        f"visible={frame.locator(FORM_ANCHOR).first.is_visible()}"
                    )
                ids = frame.locator("input:visible").evaluate_all(
                    "els => els.map(e => e.id || e.name || '').filter(Boolean).slice(0, 40)"
                )
                if ids:
                    visible_inputs.append(f"{frame.name or '(main)'}:{ids}")
            except Exception:
                continue

        raise RuntimeError(
            f"未找到可见的报关数据查询表单（iframe={IFRAME_NAME}），"
            f"候选: {candidates or '无'}；可见输入框: {visible_inputs or '无'}"
        )

    @staticmethod
    def _ensure_autocompleter(frame, hidden_id: str, text_id: str, code: str, label: str) -> None:
        """写入自动完成控件：隐藏域存编码，文本域显示名称"""
        frame.locator(f"#{hidden_id}").evaluate(JS_SET_VALUE, code)
        frame.locator(f"#{text_id}").evaluate(JS_SET_VALUE, label)

    @staticmethod
    def _set_value(frame, selector: str, value: str) -> None:
        """给日期等普通输入框赋值"""
        frame.locator(selector).evaluate(JS_SET_VALUE, value)

    @staticmethod
    def _check_value(frame, selector: str, expected: str, field_name: str) -> None:
        """校验控件值是否写入成功"""
        actual = frame.locator(selector).input_value()
        if actual != expected:
            raise RuntimeError(f"{field_name}未填写成功: 期望 {expected}，实际 {actual}")
