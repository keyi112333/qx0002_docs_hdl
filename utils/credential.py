#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Windows 凭据读取工具
"""
from log.loggingConfig import logger

try:
    import win32cred
    HAS_WIN32CRED = True
except ImportError:
    HAS_WIN32CRED = False


class CredentialError(Exception):
    """凭据读取失败"""
    pass


def get_password(target_name: str) -> str:
    """从 Windows 凭据管理器读取普通凭据的密码

    Args:
        target_name: 凭据目标名（TargetName），如 QX0002_SINGLEWINDOW_珠海

    Returns:
        str: 凭据中保存的密码

    Raises:
        CredentialError: 凭据不存在或读取失败
    """
    if not HAS_WIN32CRED:
        raise CredentialError("需要安装 pywin32: pip install pywin32")

    target = str(target_name).strip() if target_name else ""
    if not target:
        raise CredentialError("凭据目标名为空")

    try:
        credential = win32cred.CredRead(target, win32cred.CRED_TYPE_GENERIC)
    except Exception as e:
        raise CredentialError(
            f"读取 Windows 凭据失败: {target}，请在凭据管理器中维护该普通凭据（{e}）"
        ) from e

    blob = credential.get("CredentialBlob") or b""
    if isinstance(blob, str):
        password = blob
    else:
        try:
            password = blob.decode("utf-16-le")
        except UnicodeDecodeError:
            password = blob.decode("utf-8", errors="ignore")
    password = password.rstrip("\x00")

    if not password:
        raise CredentialError(f"Windows 凭据中未保存密码: {target}")

    logger.info(f"已读取 Windows 凭据: {target}")
    return password
