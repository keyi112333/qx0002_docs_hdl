#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
单一窗口独立测试脚本

只执行“单一窗口登录 + 报关数据查询 + 列表导出/详情导出/单据下载”部分：
- 不加载 UserConfig，不执行 initEnv（不会关闭其他进程）
- 公司、凭据目标名、浏览器参数直接读取 config.xlsx
- 卡口令从 Windows 凭据管理器读取
- 单一窗口打印的报关单存到临时目录（06_Temp/{公司}/YYYYMM），仅供提取信息，
  真正归档的报关单以电子口岸导出为准（test_chinaport.py）

用法：
    python tests/test_single_window.py
    python tests/test_single_window.py --company 珠海 --pause 60
    python tests/test_single_window.py --export-list
    python tests/test_single_window.py --export-list --detail-no E20260001897868651
    python tests/test_single_window.py --export-list --detail-no E20260001897868651 --bonded-no QD573026E000091160
    python tests/test_single_window.py --config "data/.../config.xlsx"
"""
import argparse
import sys
import time
from datetime import date, timedelta
from html.parser import HTMLParser
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from config.config_loader import ConfigError, ConfigLoader
from log.loggingConfig import logger
from utils.credential import CredentialError, get_password
from utils.naming import DECLARATION_DOC_TYPE, build_document_path, parse_export_date
from web.browser import BrowserManager
from web.single_window.bonded_list import BondedListQuery
from web.single_window.declaration import DeclarationExport
from web.single_window.login import SingleWindowLogin
from web.single_window.query import DeclarationQuery

DEFAULT_CONFIG = (
    ROOT / "data" / "QX0002_报关单和放行通知单等单据的处理" / "01_Config" / "config.xlsx"
)
DEFAULT_OUTPUT_DIR = ROOT / "data" / "QX0002_报关单和放行通知单等单据的处理" / "03_Output"
DEFAULT_TEMP_DIR = ROOT / "data" / "QX0002_报关单和放行通知单等单据的处理" / "06_Temp"


def _as_bool(value, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    return str(value).strip().lower() in ("true", "1", "yes", "是")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="单一窗口登录+查询独立测试")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG), help="config.xlsx 路径")
    parser.add_argument("--company", default=None, help="公司名称，缺省取第一家启用的公司")
    parser.add_argument("--days", type=int, default=7, help="最近操作时间天数，默认 7 天")
    parser.add_argument("--pause", type=int, default=30, help="查询后浏览器保持打开的秒数，0 表示立即关闭")
    parser.add_argument("--no-submit", action="store_true", help="只填写查询条件、不点击查询，便于人工确认")
    parser.add_argument("--export-list", action="store_true", help="查询后全选并导出报关单列表")
    parser.add_argument(
        "--detail-no", action="append", default=[], metavar="统一编号",
        help="按统一编号查询、详情导出并下载三个单据；不指定时取列表导出第一条",
    )
    parser.add_argument("--out-dir", default=str(DEFAULT_OUTPUT_DIR), help="导出目录，默认 data/.../03_Output")
    parser.add_argument(
        "--temp-dir", default=str(DEFAULT_TEMP_DIR),
        help="临时目录，默认 data/.../06_Temp；单一窗口打印的报关单存这里（正确版本以电子口岸导出为准）",
    )
    parser.add_argument(
        "--bonded-no", default="QD573026E000091160", metavar="保税核注清单编号",
        help="保税核注清单打印下载；留空则跳过（默认用最新录制的清单号）",
    )
    parser.add_argument(
        "--bonded-days", type=int, default=62,
        help="保税核注清单查询天数，默认当天往前 62 天（日期跨度需小于 90 天）",
    )
    parser.add_argument(
        "--contract-no", default=None,
        help="合同编号（来自报关单合同协议号）；不传用占位符“无合同号”，仅用于文件命名",
    )
    parser.add_argument(
        "--export-date", default=None,
        help="出口日期 YYYYMMDD（来自报关单详情 Excel）；不传用当天，仅用于命名与归档年月",
    )
    parser.add_argument(
        "--entry-no", default=None, metavar="报关单号",
        help="报关单号（海关编号，用于单据命名）；不传时退回统一编号",
    )
    return parser.parse_args()


def load_company(config_path: Path, company_name: str = None):
    """读取公司盾配置，返回 (配置加载器, 公司行)"""
    if not config_path.exists():
        raise ConfigError(f"config.xlsx 不存在: {config_path}")

    loader = ConfigLoader(str(config_path))
    companies = [
        company
        for company in loader.read_table("公司盾配置")
        if str(company.get("启用") or "").strip() in ("是", "true", "True", "1")
    ]
    if not companies:
        raise ConfigError("公司盾配置中没有启用的公司")

    if company_name:
        for company in companies:
            if str(company.get("公司名称") or "").strip() == company_name:
                return loader, company
        raise ConfigError(f"公司盾配置中未找到公司: {company_name}")

    return loader, companies[0]


def _read_text(list_path: Path) -> str:
    """按 UTF-8 / GB18030 依次尝试解码"""
    raw = list_path.read_bytes()
    for encoding in ("utf-8", "gb18030"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


class _TableRowParser(HTMLParser):
    """解析 HTML 表格，按行返回单元格文本（单一窗口的 xls 导出实为 HTML）"""

    def __init__(self):
        super().__init__()
        self.rows = []
        self._row = None
        self._cell = None

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._row = []
        elif tag in ("td", "th") and self._row is not None:
            self._cell = []

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._cell is not None:
            self._row.append("".join(self._cell).strip())
            self._cell = None
        elif tag == "tr" and self._row is not None:
            self.rows.append(self._row)
            self._row = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)


def _iter_excel_rows(list_path: Path):
    """按实际内容读取导出文件的行（xlsx 用 openpyxl，xls 用 xlrd 或 HTML 解析）"""
    if list_path.suffix.lower() == ".xls":
        with list_path.open("rb") as file:
            head = file.read(4)
        if head.startswith(b"<"):
            parser = _TableRowParser()
            parser.feed(_read_text(list_path))
            yield from parser.rows
            return
        try:
            import xlrd
        except ImportError:
            raise RuntimeError("导出文件为二进制 .xls，请先安装 xlrd: pip install xlrd") from None
        sheet = xlrd.open_workbook(str(list_path)).sheet_by_index(0)
        for index in range(sheet.nrows):
            yield sheet.row_values(index)
        return

    workbook = openpyxl.load_workbook(list_path, data_only=True, read_only=True)
    try:
        for row in workbook.active.iter_rows(values_only=True):
            yield row
    finally:
        workbook.close()


def read_first_unified_no(list_path: Path) -> str:
    """读取报关单列表导出文件中的第一条统一编号"""
    column = None
    for row in _iter_excel_rows(list_path):
        values = ["" if value is None else str(value).strip() for value in row]
        if column is None:
            if "统一编号" in values:
                column = values.index("统一编号")
            continue
        if column < len(values) and values[column]:
            return values[column]
    raise ValueError(f"导出列表中没有找到统一编号: {list_path}")


def run(args: argparse.Namespace) -> None:
    loader, company = load_company(Path(args.config), args.company)
    company_name = str(company.get("公司名称") or "").strip()
    credential_target = str(company.get("凭据目标名") or "").strip()
    if not credential_target:
        raise ConfigError(f"公司未配置凭据目标名: {company_name}")

    logger.info(f"测试公司: {company_name}，凭据目标名: {credential_target}")
    password = get_password(credential_target)

    headless = _as_bool(loader.get("系统配置.无头模式", False))
    channel = loader.get("系统配置.浏览器通道", "msedge")
    url = loader.get("系统配置.单一窗口地址", "https://www.singlewindow.cn/#/")
    page_timeout = int(loader.get("系统配置.页面超时(秒)", 30)) * 1000
    login_timeout = int(loader.get("系统配置.登录超时(秒)", 120)) * 1000

    browser = BrowserManager(headless=headless, channel=channel)
    browser.start()
    try:
        login = SingleWindowLogin(browser, url=url, page_timeout=page_timeout)
        login.login(password, timeout=login_timeout)
        goods_page = login.open_declaration_query()

        end_date = date.today()
        start_date = end_date - timedelta(days=args.days - 1)
        DeclarationQuery(goods_page).search(start_date, end_date, submit=not args.no_submit)

        exported = []
        stamp = time.strftime("%Y%m%d %H%M%S")
        out_dir = Path(args.out_dir)
        download_timeout = int(loader.get("系统配置.下载超时(秒)", 60)) * 1000
        # 报关单号/合同编号/出口日期最终由报关单详情 Excel 提供（解析未实现），测试用命令行参数
        contract_no = (args.contract_no or "").strip()
        export_date = parse_export_date(args.export_date) if args.export_date else date.today()
        if not (args.entry_no and args.contract_no and args.export_date):
            logger.warning(
                f"报关单号/合同编号/出口日期未全部提供，按 {args.entry_no or '统一编号'}"
                f"/{contract_no or '无合同号'}/{export_date:%Y%m%d} 命名，仅用于验证下载"
            )

        if args.export_list or args.detail_no:
            if args.no_submit:
                raise ConfigError("--no-submit 只填写查询条件，不能导出，请去掉该参数")
            exporter = DeclarationExport(
                goods_page, page_timeout=page_timeout, download_timeout=download_timeout
            )

            detail_nos = list(args.detail_no)
            if args.export_list:
                list_path = out_dir / "报关单列表" / f"{company_name}-报关单列表-{stamp}.xlsx"
                exported.append(exporter.export_list(list_path))
                if not detail_nos:
                    # BIP 比对未接入前，默认取导出列表的第一条统一编号
                    detail_nos = [read_first_unified_no(list_path)]
                    logger.info(f"默认取导出列表第一条统一编号: {detail_nos[0]}")

            for cus_ciq_no in detail_nos:
                exporter.query_by_no(cus_ciq_no)
                detail_path = out_dir / "报关单详情" / f"{cus_ciq_no}-报关单详情-{stamp}.xlsx"
                exported.append(exporter.export_detail(cus_ciq_no, detail_path))
                # 单一窗口打印的报关单仅供提取信息，存 06_Temp/{公司}/YYYYMM
                declaration_path = build_document_path(
                    Path(args.temp_dir) / company_name,
                    (args.entry_no or cus_ciq_no).strip(),
                    contract_no,
                    export_date,
                    DECLARATION_DOC_TYPE,
                )
                exported.extend(
                    exporter.download_documents(
                        cus_ciq_no,
                        out_dir / "下载单据",
                        declaration_path=declaration_path,
                    ).values()
                )
                exported.extend(
                    exporter.download_invoice_packing(cus_ciq_no, out_dir / "下载单据").values()
                )

        if args.bonded_no:
            bonded_end = date.today()
            bonded_days = min(max(args.bonded_days, 1), 90)
            bonded_start = bonded_end - timedelta(days=bonded_days - 1)
            bonded = BondedListQuery(
                browser.page, page_timeout=page_timeout, download_timeout=download_timeout
            )
            bonded.open()
            found = bonded.query(args.bonded_no, bonded_start, bonded_end)
            if found:
                exported.append(
                    bonded.print_list(
                        args.bonded_no,
                        out_dir / "下载单据" / f"{args.bonded_no}-保税核注清单.pdf",
                    )
                )
            else:
                logger.warning(f"保税核注清单无数据，跳过下载: {args.bonded_no}")

        action = "已填写查询条件（未提交）" if args.no_submit else "已提交查询"
        print(f"PASS: {company_name} {action}，最近操作时间 {start_date} ~ {end_date}")
        for path in exported:
            print(f"导出文件: {path}")
        if args.pause > 0:
            print(f"浏览器保持打开 {args.pause} 秒，可人工确认查询结果...")
            goods_page.wait_for_timeout(args.pause * 1000)
    finally:
        browser.close()


def main() -> int:
    args = parse_args()
    try:
        run(args)
        print("测试通过")
        return 0
    except (ConfigError, CredentialError) as e:
        print(f"测试前置条件失败: {e}")
        return 1
    except Exception as e:
        logger.exception("单一窗口测试失败")
        print(f"测试失败: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
