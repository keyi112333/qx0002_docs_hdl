#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
单一窗口-综合查询-报关数据查询
"""
import re
from datetime import date

from log.loggingConfig import logger
from web.base_page import BasePage

IFRAME_SELECTOR = 'iframe[name="iframe01"]'


class DeclarationQuery(BasePage):
    """报关数据查询条件填写与查询

    查询条件由页面隐藏域 + 文本显示域组成（如 etpsCategory1/etpsCategory1Name），
    文本显示域由 autocompleter 渲染，直接点击文本不可靠，因此按隐藏域 id 赋值。
    """

    # (隐藏域id, 文本域id, 隐藏域值, 文本显示值)
    ETPS_CATEGORY = ("etpsCategory1", "etpsCategory1Name", "C", "C-报关收发货人")
    IE_FLAG = ("cusIEFlag1", "cusIEFlag1Name", "E", "E-出口")
    TABLE_FLAG = ("tableFlag", "tableFlagName", "1", "1-是")

    def search(self, start_date: date, end_date: date) -> bool:
        """填写查询条件并点击查询

        Args:
            start_date: 最近操作时间起
            end_date: 最近操作时间止（单次查询不超过7天）
        """
        logger.info(
            f"填写报关数据查询条件: 最近操作时间 {start_date:%Y-%m-%d} ~ {end_date:%Y-%m-%d}"
        )
        frame = self.page.frame_locator(IFRAME_SELECTOR)

        self._ensure_autocompleter(frame, *self.ETPS_CATEGORY)
        self._ensure_autocompleter(frame, *self.IE_FLAG)
        self._ensure_autocompleter(frame, *self.TABLE_FLAG)

        self._set_value(frame, "#updateTime1", start_date.strftime("%Y-%m-%d"))
        self._set_value(frame, "#updateTimeEnd1", end_date.strftime("%Y-%m-%d"))

        frame.get_by_role("button", name=re.compile("查询")).first.click()
        logger.info("已提交报关数据查询")
        return True

    @staticmethod
    def _ensure_autocompleter(frame, hidden_id: str, text_id: str, code: str, label: str) -> None:
        """确保自动完成控件已选中指定值"""
        hidden = frame.locator(f"#{hidden_id}")
        if hidden.input_value() != code:
            hidden.evaluate("(el, value) => { el.value = value; }", code)

        text = frame.locator(f"#{text_id}")
        if not text.input_value():
            text.evaluate("(el, value) => { el.value = value; }", label)

    @staticmethod
    def _set_value(frame, selector: str, value: str) -> None:
        """赋值并触发 input/change，兼容 laydate 日期控件"""
        frame.locator(selector).evaluate(
            """(el, value) => {
                el.value = value;
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
            }""",
            value,
        )
