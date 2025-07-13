"""
自定义异常类

定义应用程序特定的异常类型。
"""


class BaseApplicationError(Exception):
    """应用程序基础异常类"""
    
    def __init__(self, message: str, error_code: str = None, details: dict = None):
        """
        初始化异常
        
        Args:
            message: 错误消息
            error_code: 错误代码
            details: 错误详情
        """
        super().__init__(message)
        self.message = message
        self.error_code = error_code or self.__class__.__name__
        self.details = details or {}
    
    def __str__(self):
        if self.details:
            return f"{self.message} (错误代码: {self.error_code}, 详情: {self.details})"
        return f"{self.message} (错误代码: {self.error_code})"


class ExcelProcessingError(BaseApplicationError):
    """Excel处理相关异常"""
    pass


class FileOperationError(ExcelProcessingError):
    """文件操作异常"""
    pass


class DataFormatError(ExcelProcessingError):
    """数据格式异常"""
    pass


class RankingError(BaseApplicationError):
    """排名计算相关异常"""
    pass


class ScoreCalculationError(RankingError):
    """评分计算异常"""
    pass


class ValidationError(BaseApplicationError):
    """验证相关异常"""
    pass


class ConfigurationError(BaseApplicationError):
    """配置相关异常"""
    pass


class InitializationError(BaseApplicationError):
    """初始化异常"""
    pass


class ProcessingInterruptedError(BaseApplicationError):
    """处理中断异常"""
    pass 