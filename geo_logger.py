"""
岩土工程写作系统 - 统一日志模块
Geo Engineering Writer Logger System

功能：
- 统一的日志记录接口
- 文件日志 + 控制台日志
- 自动创建日志目录
- 时间戳、模块名、级别追踪
"""

import logging
import os
from datetime import datetime
from pathlib import Path


def setup_logger(name="geo_writer", log_dir="logs", log_level=logging.DEBUG):
    """
    设置统一的日志系统
    
    参数：
        name: 日志器名称（默认 geo_writer）
        log_dir: 日志文件目录（默认 logs/）
        log_level: 日志级别（默认 DEBUG）
    
    返回：
        logging.Logger 实例
    """
    # 确保日志目录存在
    log_path = Path(log_dir)
    log_path.mkdir(exist_ok=True)
    
    # 创建或获取日志器
    logger = logging.getLogger(name)
    logger.setLevel(log_level)
    
    # 避免重复添加 handler
    if logger.handlers:
        return logger
    
    # 创建格式化器
    fmt = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(module)s - %(funcName)s:%(lineno)d - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    
    # 文件日志处理器：记录所有级别
    log_filename = log_path / f"geo_writer_{datetime.now().strftime('%Y%m%d')}.log"
    file_handler = logging.FileHandler(log_filename, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(fmt)
    
    # 控制台日志处理器：只显示 WARNING 及以上
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.WARNING)
    console_handler.setFormatter(fmt)
    
    # 添加处理器
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    # 记录初始化
    logger.info("=" * 60)
    logger.info(f"日志系统初始化完成: {name}")
    logger.info(f"日志文件: {log_filename}")
    logger.info("=" * 60)
    
    return logger


# 创建全局日志器实例
logger = setup_logger()


# 便捷函数
def log_step(step_name, details=""):
    """记录步骤开始"""
    logger.info(f"📍 [步骤] {step_name}")
    if details:
        logger.info(f"   详情: {details}")


def log_success(step_name, details=""):
    """记录步骤成功"""
    logger.info(f"✅ [成功] {step_name}")
    if details:
        logger.info(f"   详情: {details}")


def log_error(step_name, error, details=""):
    """记录错误"""
    logger.error(f"❌ [错误] {step_name}")
    logger.error(f"   错误信息: {error}")
    if details:
        logger.error(f"   详情: {details}")


def log_warning(step_name, message, details=""):
    """记录警告"""
    logger.warning(f"⚠️  [警告] {step_name}")
    logger.warning(f"   警告信息: {message}")
    if details:
        logger.warning(f"   详情: {details}")


def log_info(message, details=""):
    """记录一般信息"""
    logger.info(f"ℹ️  [信息] {message}")
    if details:
        logger.info(f"   详情: {details}")


def log_debug(message, details=""):
    """记录调试信息"""
    logger.debug(f"🔍 [调试] {message}")
    if details:
        logger.debug(f"   详情: {details}")
