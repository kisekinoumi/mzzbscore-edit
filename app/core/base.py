"""
基础抽象类和接口定义

提供所有服务类和处理器类的基础接口。
"""

from abc import ABC, abstractmethod
from typing import Any, Dict, Optional
import logging
import pandas as pd


class BaseService(ABC):
    """服务类基础抽象类"""
    
    def __init__(self, logger: Optional[logging.Logger] = None):
        """
        初始化基础服务
        
        Args:
            logger: 日志记录器，如果为None则使用默认logger
        """
        self._logger = logger or logging.getLogger(self.__class__.__name__)
        self._initialized = False
    
    @abstractmethod
    def initialize(self) -> bool:
        """
        初始化服务
        
        Returns:
            bool: 初始化是否成功
        """
        pass
    
    @property
    def logger(self) -> logging.Logger:
        """获取日志记录器"""
        return self._logger
    
    @property
    def is_initialized(self) -> bool:
        """检查服务是否已初始化"""
        return self._initialized
    
    def _set_initialized(self, status: bool = True):
        """设置初始化状态"""
        self._initialized = status


class BaseHandler(ABC):
    """处理器类基础抽象类"""
    
    def __init__(self, logger: Optional[logging.Logger] = None):
        """
        初始化基础处理器
        
        Args:
            logger: 日志记录器
        """
        self._logger = logger or logging.getLogger(self.__class__.__name__)
        self._context: Dict[str, Any] = {}
    
    @abstractmethod
    def process(self, data: Any) -> Any:
        """
        处理数据
        
        Args:
            data: 输入数据
            
        Returns:
            Any: 处理结果
        """
        pass
    
    @property
    def logger(self) -> logging.Logger:
        """获取日志记录器"""
        return self._logger
    
    @property
    def context(self) -> Dict[str, Any]:
        """获取处理上下文"""
        return self._context.copy()
    
    def set_context(self, key: str, value: Any):
        """设置上下文值"""
        self._context[key] = value
    
    def get_context(self, key: str, default: Any = None) -> Any:
        """获取上下文值"""
        return self._context.get(key, default)


class IDataProcessor(ABC):
    """数据处理器接口"""
    
    @abstractmethod
    def validate_data(self, data: pd.DataFrame) -> bool:
        """验证数据格式"""
        pass
    
    @abstractmethod
    def process_data(self, data: pd.DataFrame) -> pd.DataFrame:
        """处理数据"""
        pass


class IFileHandler(ABC):
    """文件处理器接口"""
    
    @abstractmethod
    def read_file(self, file_path: str) -> Any:
        """读取文件"""
        pass
    
    @abstractmethod
    def write_file(self, file_path: str, data: Any) -> bool:
        """写入文件"""
        pass
    
    @abstractmethod
    def validate_file(self, file_path: str) -> bool:
        """验证文件"""
        pass 