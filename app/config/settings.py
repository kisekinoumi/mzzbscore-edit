"""
应用程序设置类

管理应用程序的配置和设置。
"""

import os
from typing import Optional, Dict, Any
from app.config.constants import DEFAULT_INPUT_FILE, DEFAULT_OUTPUT_FILE_MONTHLY, DEFAULT_OUTPUT_FILE_FINAL


class Settings:
    """应用程序设置管理器"""
    
    def __init__(self, config_dict: Optional[Dict[str, Any]] = None):
        """
        初始化设置
        
        Args:
            config_dict: 配置字典，如果为None则使用默认设置
        """
        self._config = config_dict or {}
        self._load_default_settings()
    
    def _load_default_settings(self):
        """加载默认设置"""
        default_settings = {
            # 文件设置
            "input_file": self._config.get("input_file", DEFAULT_INPUT_FILE),
            "output_file_monthly": self._config.get("output_file_monthly", DEFAULT_OUTPUT_FILE_MONTHLY),
            "output_file_final": self._config.get("output_file_final", DEFAULT_OUTPUT_FILE_FINAL),
            
            # 处理设置
            "enable_logging": self._config.get("enable_logging", True),
            "log_level": self._config.get("log_level", "INFO"),
            "log_to_file": self._config.get("log_to_file", True),
            "log_directory": self._config.get("log_directory", "."),
            
            # Excel设置
            "preserve_hyperlinks": self._config.get("preserve_hyperlinks", True),
            "apply_column_styles": self._config.get("apply_column_styles", True),
            "header_row": self._config.get("header_row", 2),
            "data_start_row": self._config.get("data_start_row", 3),
            
            # 排名设置
            "calculate_comprehensive_score": self._config.get("calculate_comprehensive_score", True),
            "ranking_method": self._config.get("ranking_method", "min"),  # min, max, average, first, dense
            "exclude_invalid_scores": self._config.get("exclude_invalid_scores", True),
            
            # 验证设置
            "strict_validation": self._config.get("strict_validation", True),
            "allow_empty_scores": self._config.get("allow_empty_scores", True),
            "check_data_integrity": self._config.get("check_data_integrity", True),
            
            # 性能设置
            "chunk_size": self._config.get("chunk_size", 1000),
            "memory_limit_mb": self._config.get("memory_limit_mb", 500),
            "enable_parallel_processing": self._config.get("enable_parallel_processing", False),
            
            # 用户界面设置
            "show_progress": self._config.get("show_progress", True),
            "max_input_attempts": self._config.get("max_input_attempts", 3),
            "auto_backup": self._config.get("auto_backup", True),
        }
        
        # 合并设置
        for key, value in default_settings.items():
            if key not in self._config:
                self._config[key] = value
    
    def get(self, key: str, default: Any = None) -> Any:
        """
        获取设置值
        
        Args:
            key: 设置键
            default: 默认值
            
        Returns:
            Any: 设置值
        """
        return self._config.get(key, default)
    
    def set(self, key: str, value: Any):
        """
        设置值
        
        Args:
            key: 设置键
            value: 设置值
        """
        self._config[key] = value
    
    def update(self, config_dict: Dict[str, Any]):
        """
        批量更新设置
        
        Args:
            config_dict: 配置字典
        """
        self._config.update(config_dict)
    
    @property
    def input_file(self) -> str:
        """输入文件路径"""
        return self.get("input_file")
    
    @property
    def output_file_monthly(self) -> str:
        """首月评分输出文件路径"""
        return self.get("output_file_monthly")
    
    @property
    def output_file_final(self) -> str:
        """完结评分输出文件路径"""
        return self.get("output_file_final")
    
    @property
    def enable_logging(self) -> bool:
        """是否启用日志"""
        return self.get("enable_logging", True)
    
    @property
    def log_level(self) -> str:
        """日志级别"""
        return self.get("log_level", "INFO")
    
    @property
    def preserve_hyperlinks(self) -> bool:
        """是否保留超链接"""
        return self.get("preserve_hyperlinks", True)
    
    @property
    def apply_column_styles(self) -> bool:
        """是否应用列样式"""
        return self.get("apply_column_styles", True)
    
    @property
    def strict_validation(self) -> bool:
        """是否启用严格验证"""
        return self.get("strict_validation", True)
    
    @property
    def show_progress(self) -> bool:
        """是否显示进度"""
        return self.get("show_progress", True)
    
    def to_dict(self) -> Dict[str, Any]:
        """
        转换为字典
        
        Returns:
            Dict[str, Any]: 配置字典
        """
        return self._config.copy()
    
    @classmethod
    def from_file(cls, file_path: str) -> 'Settings':
        """
        从文件加载设置
        
        Args:
            file_path: 配置文件路径
            
        Returns:
            Settings: 设置实例
        """
        config_dict = {}
        if os.path.exists(file_path):
            try:
                import json
                with open(file_path, 'r', encoding='utf-8') as f:
                    config_dict = json.load(f)
            except Exception:
                # 如果文件读取失败，使用默认设置
                pass
        
        return cls(config_dict)
    
    def save_to_file(self, file_path: str) -> bool:
        """
        保存设置到文件
        
        Args:
            file_path: 配置文件路径
            
        Returns:
            bool: 是否保存成功
        """
        try:
            import json
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            return False 