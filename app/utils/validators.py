"""
验证工具类

提供文件、数据等验证功能。
"""

import os
import pandas as pd
from pathlib import Path
from typing import List, Optional
from app.utils.exceptions import ValidationError, FileOperationError


class FileValidator:
    """文件验证器"""
    
    SUPPORTED_EXCEL_EXTENSIONS = ['.xlsx', '.xls']
    
    @staticmethod
    def validate_file_exists(file_path: str) -> bool:
        """
        验证文件是否存在
        
        Args:
            file_path: 文件路径
            
        Returns:
            bool: 文件是否存在
            
        Raises:
            FileOperationError: 文件不存在
        """
        if not os.path.exists(file_path):
            raise FileOperationError(f"文件不存在: {file_path}")
        return True
    
    @staticmethod
    def validate_file_readable(file_path: str) -> bool:
        """
        验证文件是否可读
        
        Args:
            file_path: 文件路径
            
        Returns:
            bool: 文件是否可读
            
        Raises:
            FileOperationError: 文件无法读取
        """
        if not os.access(file_path, os.R_OK):
            raise FileOperationError(f"文件无读取权限: {file_path}")
        return True
    
    @staticmethod
    def validate_excel_file(file_path: str) -> bool:
        """
        验证是否为支持的Excel文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            bool: 是否为支持的Excel文件
            
        Raises:
            ValidationError: 不支持的文件格式
        """
        file_extension = Path(file_path).suffix.lower()
        if file_extension not in FileValidator.SUPPORTED_EXCEL_EXTENSIONS:
            raise ValidationError(
                f"不支持的文件格式: {file_extension}",
                details={"supported_formats": FileValidator.SUPPORTED_EXCEL_EXTENSIONS}
            )
        return True
    
    @staticmethod
    def validate_output_directory(file_path: str) -> bool:
        """
        验证输出目录是否可写
        
        Args:
            file_path: 输出文件路径
            
        Returns:
            bool: 目录是否可写
            
        Raises:
            FileOperationError: 目录无写入权限
        """
        output_dir = os.path.dirname(file_path) or "."
        
        # 如果目录不存在，尝试创建
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir, exist_ok=True)
            except OSError as e:
                raise FileOperationError(f"无法创建输出目录 {output_dir}: {e}")
        
        # 检查写入权限
        if not os.access(output_dir, os.W_OK):
            raise FileOperationError(f"输出目录没有写入权限: {output_dir}")
        
        return True
    
    @classmethod
    def validate_complete_file_operation(cls, input_file: str, output_file: str) -> bool:
        """
        完整的文件操作验证
        
        Args:
            input_file: 输入文件路径
            output_file: 输出文件路径
            
        Returns:
            bool: 验证是否通过
        """
        cls.validate_file_exists(input_file)
        cls.validate_file_readable(input_file)
        cls.validate_excel_file(input_file)
        cls.validate_output_directory(output_file)
        return True


class DataValidator:
    """数据验证器"""
    
    @staticmethod
    def validate_dataframe_not_empty(df: pd.DataFrame, name: str = "DataFrame") -> bool:
        """
        验证DataFrame不为空
        
        Args:
            df: 要验证的DataFrame
            name: DataFrame名称（用于错误消息）
            
        Returns:
            bool: DataFrame是否有效
            
        Raises:
            ValidationError: DataFrame为空或None
        """
        if df is None:
            raise ValidationError(f"{name}不能为None")
        
        if df.empty:
            raise ValidationError(f"{name}不能为空")
        
        return True
    
    @staticmethod
    def validate_required_columns(df: pd.DataFrame, required_columns: List[str]) -> bool:
        """
        验证DataFrame包含必需的列
        
        Args:
            df: 要验证的DataFrame
            required_columns: 必需的列名列表
            
        Returns:
            bool: 是否包含所有必需列
            
        Raises:
            ValidationError: 缺少必需的列
        """
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValidationError(
                f"缺少必需的列: {missing_columns}",
                details={
                    "missing_columns": missing_columns,
                    "available_columns": list(df.columns),
                    "required_columns": required_columns
                }
            )
        return True
    
    @staticmethod
    def validate_numeric_column(df: pd.DataFrame, column: str, allow_nan: bool = True) -> bool:
        """
        验证列是否包含有效的数值数据
        
        Args:
            df: 要验证的DataFrame
            column: 列名
            allow_nan: 是否允许NaN值
            
        Returns:
            bool: 列是否有效
            
        Raises:
            ValidationError: 列包含无效数据
        """
        if column not in df.columns:
            raise ValidationError(f"列 '{column}' 不存在")
        
        series = df[column]
        
        # 检查是否有有效的数值
        numeric_series = pd.to_numeric(series, errors='coerce')
        valid_count = numeric_series.notna().sum()
        
        if valid_count == 0:
            raise ValidationError(
                f"列 '{column}' 中没有有效的数值数据",
                details={"column": column, "total_rows": len(series)}
            )
        
        if not allow_nan and series.isna().any():
            nan_count = series.isna().sum()
            raise ValidationError(
                f"列 '{column}' 包含 {nan_count} 个空值，但不允许空值",
                details={"column": column, "nan_count": nan_count}
            )
        
        return True
    
    @staticmethod
    def validate_data_integrity(df: pd.DataFrame, key_column: str) -> bool:
        """
        验证数据完整性
        
        Args:
            df: 要验证的DataFrame
            key_column: 主键列名
            
        Returns:
            bool: 数据是否完整
            
        Raises:
            ValidationError: 数据完整性问题
        """
        # 检查主键列是否存在
        if key_column not in df.columns:
            raise ValidationError(f"主键列 '{key_column}' 不存在")
        
        # 检查主键列是否有空值
        null_count = df[key_column].isnull().sum()
        if null_count > 0:
            raise ValidationError(
                f"主键列 '{key_column}' 包含 {null_count} 个空值",
                details={"column": key_column, "null_count": null_count}
            )
        
        # 检查是否有重复值
        duplicate_count = df[key_column].duplicated().sum()
        if duplicate_count > 0:
            raise ValidationError(
                f"主键列 '{key_column}' 包含 {duplicate_count} 个重复值",
                details={"column": key_column, "duplicate_count": duplicate_count}
            )
        
        return True 