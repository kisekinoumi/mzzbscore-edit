"""
统一的日志管理类

提供应用程序的日志记录功能。
"""

import logging
import os
import sys
from datetime import datetime
from typing import Optional


class Logger:
    """统一日志管理器"""
    
    _instance: Optional['Logger'] = None
    _logger: Optional[logging.Logger] = None
    
    def __new__(cls):
        """单例模式"""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        """初始化日志管理器"""
        if self._logger is None:
            self._setup_logger()
    
    def _setup_logger(self):
        """设置日志记录器"""
        try:
            # 创建日志记录器
            self._logger = logging.getLogger('mzzbscore')
            self._logger.setLevel(logging.INFO)
            
            # 清除已有的处理器
            self._logger.handlers.clear()
            
            # 设置日志格式
            formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                datefmt='%Y-%m-%d %H:%M:%S'
            )
            
            # 控制台处理器
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setLevel(logging.INFO)
            console_handler.setFormatter(formatter)
            self._logger.addHandler(console_handler)
            
            # 文件处理器
            try:
                # 直接在根目录创建日志文件
                log_file = f"processing_{datetime.now().strftime('%Y%m%d')}.log"
                file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
                file_handler.setLevel(logging.DEBUG)
                file_handler.setFormatter(formatter)
                self._logger.addHandler(file_handler)
                
            except (OSError, PermissionError) as e:
                # 如果无法创建文件日志，只使用控制台日志
                self._logger.warning(f"无法创建文件日志: {e}")
            
            self._logger.info("日志系统初始化成功")
            
        except Exception as e:
            # 如果日志设置失败，使用基本配置
            logging.basicConfig(level=logging.INFO)
            self._logger = logging.getLogger('mzzbscore')
            self._logger.error(f"日志系统初始化失败，使用基本配置: {e}")
    
    @property
    def logger(self) -> logging.Logger:
        """获取日志记录器"""
        return self._logger
    
    @classmethod
    def get_logger(cls, name: str = None) -> logging.Logger:
        """
        获取日志记录器实例
        
        Args:
            name: 日志记录器名称，如果为None则使用主logger
            
        Returns:
            logging.Logger: 日志记录器
        """
        instance = cls()
        if name:
            return instance.logger.getChild(name)
        return instance.logger
    
    def set_level(self, level: int):
        """
        设置日志级别
        
        Args:
            level: 日志级别
        """
        self._logger.setLevel(level)
        for handler in self._logger.handlers:
            handler.setLevel(level)
    
    def info(self, message: str):
        """记录信息日志"""
        self._logger.info(message)
    
    def debug(self, message: str):
        """记录调试日志"""
        self._logger.debug(message)
    
    def warning(self, message: str):
        """记录警告日志"""
        self._logger.warning(message)
    
    def error(self, message: str):
        """记录错误日志"""
        self._logger.error(message)
    
    def critical(self, message: str):
        """记录严重错误日志"""
        self._logger.critical(message) 