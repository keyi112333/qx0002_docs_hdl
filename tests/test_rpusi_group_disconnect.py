#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
RPUSI NetworkUSB 分组断开连接独立测试脚本

只验证“断开指定设备的指定 USB 端口组”这一条命令（rfsnu -d）：
- 不启动浏览器、不加载 config.xlsx、不影响主流程
- 设备地址、组号需要人工提供，脚本内留空；也可用命令行参数临时传入
- 组断开后，该组可被其他主机使用；组内所有端口一起断开，不能只断开单个端口

命令依据：《NetworkUSB 远程 USB 服务器命令行操作手册》
    rfsnu -d <device address> <groupnumber>    断开指定设备的组
    rfsnu -c <device address> <groupnumber>    连接指定设备的组（对照组，供手工验证用）
    rfsnu -L                                   列表当前添加的设备信息（本脚本执行前后各查一次）

device address 格式：(IP 或 域名)[:TCP端口]，TCP 端口可省略，出厂默认 3240。

用法：
    python tests/test_rpusi_group_disconnect.py
    python tests/test_rpusi_group_disconnect.py --ip 192.168.1.250 --group 1
    python tests/test_rpusi_group_disconnect.py --ip 192.168.1.250:3240 --group 1 --dry-run
    python tests/test_rpusi_group_disconnect.py --rfsnu "C:\\Program Files\\RFS NetworkUSB\\rfsnu.exe"
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from log.loggingConfig import logger

# ===================== 需要人工提供的信息（留空，运行时补全） =====================
# TODO: 设备地址，格式 "IP" 或 "IP:TCP端口"，例如 "192.168.1.250" / "192.168.1.250:3240"
DEVICE_ADDRESS = ""

# TODO: 要断开的组号（整数，从 1 开始），例如 1
GROUP_NUMBER = None

# TODO: rfsnu.exe 完整路径；留空则在常见安装目录自动查找
RFS_NU_PATH = ""
# ==============================================================================

# rfsnu.exe 一般随 RFS NetworkUSB 客户端安装在 Program Files 下的相关目录
SEARCH_ROOTS = [
    Path(os.environ.get("ProgramFiles", r"C:\Program Files")),
    Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")),
]
SEARCH_DIR_PATTERN = re.compile(r"rfs|networkusb|rpusi", re.IGNORECASE)
EXE_NAMES = ("rfsnu.exe", "RfsNu.exe", "rfsnu")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="RPUSI NetworkUSB 分组断开连接测试")
    parser.add_argument("--ip", default=DEVICE_ADDRESS, help="设备地址，如 192.168.1.250 或 192.168.1.250:3240")
    parser.add_argument("--group", type=int, default=GROUP_NUMBER, help="要断开的组号，例如 1")
    parser.add_argument("--rfsnu", default=RFS_NU_PATH, help="rfsnu.exe 路径，缺省自动查找")
    parser.add_argument("--timeout", type=int, default=20, help="单条命令超时秒数，默认 20")
    parser.add_argument("--dry-run", action="store_true", help="只打印将执行的命令，不真正执行")
    parser.add_argument("--skip-list", action="store_true", help="跳过执行前后的 rfsnu -L 设备列表检查")
    return parser.parse_args()


def find_rfsnu(explicit_path: str) -> str:
    """查找 rfsnu.exe：优先使用显式路径，其次在 Program Files 相关目录，最后看 PATH"""
    if explicit_path:
        path = Path(explicit_path)
        if path.is_file():
            return str(path)
        raise FileNotFoundError(f"指定的 rfsnu 不存在: {explicit_path}")

    for root in SEARCH_ROOTS:
        if not root.exists():
            continue
        for sub in root.iterdir():
            if not sub.is_dir() or not SEARCH_DIR_PATTERN.search(sub.name):
                continue
            for exe in sub.rglob("rfsnu.exe"):
                return str(exe)

    for name in EXE_NAMES:
        found = shutil.which(name)
        if found:
            return found
    return ""


def _decode(data: bytes) -> str:
    """rfsnu 输出多为中文，Windows 下按 GBK 解码，失败时回退"""
    for encoding in ("gbk", "utf-8"):
        try:
            return data.decode(encoding).strip()
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace").strip()


def run_rfsnu(exe: str, args: list, timeout: int) -> subprocess.CompletedProcess:
    logger.info("执行命令: %s %s", exe, " ".join(args))
    try:
        proc = subprocess.run([exe, *args], capture_output=True, timeout=timeout)
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(f"命令超时（{timeout}s）: {exe} {' '.join(args)}") from e

    stdout = _decode(proc.stdout or b"")
    stderr = _decode(proc.stderr or b"")
    if stdout:
        logger.info("标准输出: %s", stdout)
    if stderr:
        logger.warning("标准错误: %s", stderr)
    logger.info("返回码: %s", proc.returncode)
    return proc


def list_devices(exe: str, timeout: int) -> None:
    """rfsnu -L：列出当前已添加的设备，用于执行前后对比"""
    try:
        run_rfsnu(exe, ["-L"], timeout)
    except RuntimeError as e:
        logger.warning(f"设备列表查询失败（不影响断开操作）: {e}")


def main() -> int:
    args = parse_args()

    address = str(args.ip or "").strip()
    group = args.group
    if not address:
        logger.error(
            "缺少设备地址：请填写脚本顶部 DEVICE_ADDRESS，或使用 --ip 传入，"
            "例如 --ip 192.168.1.250 或 --ip 192.168.1.250:3240"
        )
        return 2
    if group is None or int(group) < 1:
        logger.error(
            "缺少组号：请填写脚本顶部 GROUP_NUMBER，或使用 --group 传入（组号从 1 开始）"
        )
        return 2
    group = int(group)

    try:
        exe = find_rfsnu(args.rfsnu)
    except FileNotFoundError as e:
        logger.error(str(e))
        return 2
    if not exe:
        if not args.dry_run:
            logger.error(
                "未找到 rfsnu.exe：请先安装 RFS NetworkUSB 客户端（同目录安装包 NetworkUSB20250930_SHC.exe），"
                "或使用 --rfsnu 指定完整路径"
            )
            return 2
        exe = "<未找到 rfsnu.exe，请用 --rfsnu 指定>"

    logger.info("设备地址: %s，组号: %s，rfsnu: %s", address, group, exe)

    if args.dry_run:
        logger.info("dry-run：将执行 %s -L 和 %s -d %s %s", exe, exe, address, group)
        return 0

    if not args.skip_list:
        logger.info("---- 断开前设备列表 ----")
        list_devices(exe, args.timeout)

    # 断开指定设备的指定组
    proc = run_rfsnu(exe, ["-d", address, str(group)], args.timeout)
    if proc.returncode != 0:
        logger.error(
            "断开命令返回非 0（%s）：请检查设备地址是否可达、TCP 3240 是否放行、"
            "设备是否已在客户端添加、组号是否正确",
            proc.returncode,
        )
        return proc.returncode

    logger.info("组断开命令已下发：%s 第 %s 组", address, group)

    if not args.skip_list:
        logger.info("---- 断开后设备列表 ----")
        list_devices(exe, args.timeout)

    logger.info(
        "请人工确认：客户端中该组颜色由红色/橙色变回灰色、组内 USB Port 由蓝色变回绿色；"
        "如需恢复连接可执行 %s -c %s %s",
        exe,
        address,
        group,
    )
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:  # 独立测试脚本，异常直接打日志便于手工排查
        logger.exception(f"测试执行异常: {exc}")
        sys.exit(1)
