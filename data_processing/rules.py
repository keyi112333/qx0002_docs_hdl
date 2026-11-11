#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
业务规则：时间分段、待处理清单、单据完整性、状态取值（开发文档 8.4 / 8.5 / 8.6 / 8.10 / 9.1 / 9.5）
"""
from datetime import date, timedelta
from typing import Dict, Iterable, List, Tuple

from log.loggingConfig import logger

# 运行状态（8.10.1）
PENDING = "待处理"
SKIP_UPLOADED = "跳过-已上传附件"
SKIP_SUCCESS = "跳过-已成功"
DOWNLOADING = "下载中"
DOWNLOADED = "下载完成"
ENTERING = "录入中"
ENTERED = "录入完成"
BINDING = "绑定中"
SUCCESS = "成功"
SUCCESS_MANUAL = "成功-人工已处理"
TO_RERUN = "待补跑"
MANUAL = "需人工"
FAIL_DOWNLOAD = "失败-下载失败"
FAIL_ENTRY = "失败-录入失败"
FAIL_PARSE = "失败-解析失败"
FAIL_LOCKED = "失败-文件占用"

# 补跑范围：待补跑/失败（8.11）
RERUN_STATUSES = (TO_RERUN, FAIL_DOWNLOAD, FAIL_ENTRY, FAIL_LOCKED)

# 六类交付单据（4.3）
DOC_DECLARATION = "报关单"
DOC_RELEASE = "放行通知单"
DOC_ENTRUST = "电子委托协议"
DOC_INVOICE = "发票"
DOC_PACKING = "装箱单"
DOC_BONDED = "保税核注清单"

# 结果表“文件路径”列名（第 19～24 列，模板列名原样）
PATH_COLUMNS = {
    DOC_RELEASE: "放行通知单路径",
    DOC_ENTRUST: "电子委托协议路径",
    DOC_INVOICE: "发票路径",
    DOC_PACKING: "装箱单路径",
    DOC_BONDED: "报税核注清单路径",
    DOC_DECLARATION: "报关单路径",
}


def build_windows(today: date, total_days: int = 62, step: int = 7) -> List[Tuple[date, date]]:
    """按“今天向前、每段 step 天”拆分查询窗口（8.4）"""
    windows, end, remain = [], today, total_days
    while remain > 0:
        span = min(step, remain)
        start = end - timedelta(days=span - 1)
        windows.append((start, end))
        end = start - timedelta(days=1)
        remain -= span
    return windows


def need_bonded(supervision_mode: str, keyword: str = "进料") -> bool:
    """监管方式（中文名）包含配置关键字时需下载保税核注清单（9.5）"""
    return bool(keyword) and keyword in (supervision_mode or "")


def required_docs(bonded_required: bool) -> Tuple[str, ...]:
    """必需单据：报关单、放行通知单、电子委托协议、发票、装箱单；保税核注清单按监管方式"""
    docs = [DOC_DECLARATION, DOC_RELEASE, DOC_ENTRUST, DOC_INVOICE, DOC_PACKING]
    if bonded_required:
        docs.append(DOC_BONDED)
    return tuple(docs)


def dedupe_by_no(items: Iterable) -> List:
    """按报关单号去重，保留最早出现的一条（8.11）"""
    seen, result = set(), []
    for item in items:
        if item.no in seen:
            continue
        seen.add(item.no)
        result.append(item)
    return result


def build_todo(list_items: Iterable, bip_data, result_status: Dict[str, str],
               rerun_enabled: bool = True) -> Tuple[List, Dict[str, int]]:
    """生成待处理清单（8.5）

    待处理清单 = 本次查询结果 ∩ BIP 无附件记录 - 结果表已成功/需人工 + 补跑清单

    Args:
        list_items: 本次合并后的报关单列表（含统一编号、海关编号、监管方式等）
        bip_data: BipAttachmentData
        result_status: {报关单号: 运行状态}，覆盖今年与去年两个年度
        rerun_enabled: 是否把结果表中“待补跑/失败”的报关单重新纳入
    """
    items = [item for item in list_items if item.is_export_closed()]
    no_attachment = bip_data.bills_without_attachment()

    stats = {
        "本次查询数": len(items),
        "BIP无附件单号数": len(no_attachment),
        "交集待处理数": 0,
        "已成功跳过数": 0,
        "需人工跳过数": 0,
        "补跑数": 0,
        "补跑缺列表记录数": 0,
    }

    candidates = []
    for item in items:
        if item.no not in no_attachment:
            continue
        stats["交集待处理数"] += 1
        status = result_status.get(item.no, "")
        if status in (SUCCESS, SUCCESS_MANUAL):
            stats["已成功跳过数"] += 1
            continue
        if status == MANUAL:
            stats["需人工跳过数"] += 1
            continue
        candidates.append(item)

    todo = list(candidates)
    if rerun_enabled:
        by_no = {item.no: item for item in items}
        for bill_no, status in result_status.items():
            if status not in RERUN_STATUSES:
                continue
            item = by_no.get(bill_no)
            if item is None:
                # 结果表有记录但本次查询窗口未查到，缺统一编号无法逐单查询，留待下次
                stats["补跑缺列表记录数"] += 1
                logger.warning(f"补跑单不在本次查询结果中，跳过: {bill_no}（状态={status}）")
                continue
            stats["补跑数"] += 1
            todo.append(item)

    todo = dedupe_by_no(todo)
    stats["最终待处理数"] = len(todo)
    logger.info(
        "待处理清单: " + "，".join(f"{name}={value}" for name, value in stats.items())
    )
    return todo, stats


def result_row_from_list_item(item) -> dict:
    """列表导出可直接提供的初始结果字段（详情解析后再补其余列）"""
    return {
        "销售组织": item.domestic_consignor,
        "贸易种类": item.supervision_mode,
        "合同号": item.contract_no,
        "出口日期": item.export_date.strftime("%Y-%m-%d") if item.export_date else "",
        "监管方式": item.supervision_mode,
    }


def result_row_from_declaration(declaration) -> dict:
    """详情解析后可写入的 BIP/报关单字段（第 1～18 列，报关汇率待 BIP 阶段补）"""
    unit_price = declaration.unit_price()
    return {
        "销售组织": declaration.domestic_consignor,
        "贸易种类": declaration.supervision_mode or declaration.supervision_code,
        "报关客户": declaration.overseas_consignor,
        "报关币种": declaration.currency,
        "合同号": declaration.contract_no,
        "报关数量": float(declaration.total_qty),
        "报关单价": float(unit_price) if unit_price is not None else "",
        "总价": float(declaration.total_amount),
        "运输费": declaration.freight_flag,
        "保险费": declaration.insurance_flag,
        "出口日期": declaration.export_date.strftime("%Y-%m-%d"),
        "手册号": declaration.record_no,
        "贸易方式": "出口",
        "监管方式": declaration.supervision_mode or declaration.supervision_code,
        "成交方式": declaration.trade_mode_code,
        "保税核注清单编号": declaration.bonded_no,
    }
