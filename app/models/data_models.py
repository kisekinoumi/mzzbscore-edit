"""
数据模型类

定义应用程序中使用的数据结构。
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Any, Union
import pandas as pd


@dataclass
class AnimeData:
    """动漫数据模型"""
    
    original_name: str
    translated_name: Optional[str] = None
    bangumi_score: Optional[float] = None
    bangumi_total: Optional[int] = None
    anilist_score: Optional[float] = None
    anilist_total: Optional[int] = None
    myanimelist_score: Optional[float] = None
    myanimelist_total: Optional[int] = None
    filmarks_score: Optional[float] = None
    filmarks_total: Optional[int] = None
    comprehensive_score: Optional[float] = None
    notes: Optional[str] = None
    x_score: Optional[float] = None
    x_fan: Optional[str] = None
    bangumi_url: Optional[str] = None
    anilist_url: Optional[str] = None
    myanilist_url: Optional[str] = None
    filmarks_url: Optional[str] = None
    
    # 排名字段
    bangumi_rank: Optional[int] = None
    anilist_rank: Optional[int] = None
    myanimelist_rank: Optional[int] = None
    filmarks_rank: Optional[int] = None
    comprehensive_rank: Optional[int] = None
    
    def __post_init__(self):
        """数据验证和清理"""
        if not self.original_name or not self.original_name.strip():
            raise ValueError("原名不能为空")
        
        self.original_name = self.original_name.strip()
        if self.translated_name:
            self.translated_name = self.translated_name.strip()
    
    @property
    def has_valid_scores(self) -> bool:
        """检查是否有有效的评分数据"""
        scores = [self.bangumi_score, self.anilist_score, self.myanimelist_score, self.filmarks_score]
        return any(score is not None and score > 0 for score in scores)
    
    @property
    def valid_scores(self) -> Dict[str, float]:
        """获取所有有效的评分"""
        scores = {}
        if self.bangumi_score is not None and self.bangumi_score > 0:
            scores['bangumi'] = self.bangumi_score
        if self.anilist_score is not None and self.anilist_score > 0:
            scores['anilist'] = self.anilist_score
        if self.myanimelist_score is not None and self.myanimelist_score > 0:
            scores['myanimelist'] = self.myanimelist_score
        if self.filmarks_score is not None and self.filmarks_score > 0:
            scores['filmarks'] = self.filmarks_score
        return scores
    
    @property
    def should_exclude_from_ranking(self) -> bool:
        """检查是否应该从排名中排除"""
        exclude_notes = ["*时长不足", "*数据不足"]
        return self.notes in exclude_notes if self.notes else False
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典格式"""
        return {
            "原名": self.original_name,
            "译名": self.translated_name,
            "Bangumi": self.bangumi_score,
            "Bangumi_total": self.bangumi_total,
            "Bangumi_Rank": self.bangumi_rank,
            "Anilist": self.anilist_score,
            "Anilist_total": self.anilist_total,
            "Anilist_Rank": self.anilist_rank,
            "MyAnimelist": self.myanimelist_score,
            "MyAnimelist_total": self.myanimelist_total,
            "Myanimelist_Rank": self.myanimelist_rank,
            "Filmarks": self.filmarks_score,
            "Filmarks_total": self.filmarks_total,
            "Filmarks_Rank": self.filmarks_rank,
            "综合评分": self.comprehensive_score,
            "排名": self.comprehensive_rank,
            "Notes": self.notes,
            "X": self.x_score,
            "X_fan": self.x_fan,
            "Bangumi_url": self.bangumi_url,
            "Anilist_url": self.anilist_url,
            "Myanilist_url": self.myanilist_url,
            "Filmarks_url": self.filmarks_url,
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'AnimeData':
        """从字典创建实例"""
        return cls(
            original_name=data.get("原名", ""),
            translated_name=data.get("译名"),
            bangumi_score=data.get("Bangumi"),
            bangumi_total=data.get("Bangumi_total"),
            bangumi_rank=data.get("Bangumi_Rank"),
            anilist_score=data.get("Anilist"),
            anilist_total=data.get("Anilist_total"),
            anilist_rank=data.get("Anilist_Rank"),
            myanimelist_score=data.get("MyAnimelist"),
            myanimelist_total=data.get("MyAnimelist_total"),
            myanimelist_rank=data.get("Myanimelist_Rank"),
            filmarks_score=data.get("Filmarks"),
            filmarks_total=data.get("Filmarks_total"),
            filmarks_rank=data.get("Filmarks_Rank"),
            comprehensive_score=data.get("综合评分"),
            comprehensive_rank=data.get("排名"),
            notes=data.get("Notes"),
            x_score=data.get("X"),
            x_fan=data.get("X_fan"),
            bangumi_url=data.get("Bangumi_url"),
            anilist_url=data.get("Anilist_url"),
            myanilist_url=data.get("Myanilist_url"),
            filmarks_url=data.get("Filmarks_url"),
        )


@dataclass
class RankingResult:
    """排名结果模型"""
    
    valid_data: pd.DataFrame
    excluded_data: pd.DataFrame
    total_processed: int = 0
    total_valid: int = 0
    total_excluded: int = 0
    processing_time: float = 0.0
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    
    def __post_init__(self):
        """计算统计信息"""
        if self.valid_data is not None:
            self.total_valid = len(self.valid_data)
        if self.excluded_data is not None:
            self.total_excluded = len(self.excluded_data)
        self.total_processed = self.total_valid + self.total_excluded
    
    @property
    def success_rate(self) -> float:
        """成功率"""
        if self.total_processed == 0:
            return 0.0
        return self.total_valid / self.total_processed
    
    @property
    def has_errors(self) -> bool:
        """是否有错误"""
        return len(self.errors) > 0
    
    @property
    def has_warnings(self) -> bool:
        """是否有警告"""
        return len(self.warnings) > 0
    
    def add_error(self, error: str):
        """添加错误"""
        self.errors.append(error)
    
    def add_warning(self, warning: str):
        """添加警告"""
        self.warnings.append(warning)
    
    def get_summary(self) -> Dict[str, Any]:
        """获取处理摘要"""
        return {
            "total_processed": self.total_processed,
            "total_valid": self.total_valid,
            "total_excluded": self.total_excluded,
            "success_rate": f"{self.success_rate:.2%}",
            "processing_time": f"{self.processing_time:.2f}s",
            "error_count": len(self.errors),
            "warning_count": len(self.warnings),
            "has_errors": self.has_errors,
            "has_warnings": self.has_warnings
        }


@dataclass
class ProcessingConfig:
    """处理配置模型"""
    
    input_file: str
    output_file: str
    operation_type: str  # "monthly" or "final"
    preserve_hyperlinks: bool = True
    apply_styles: bool = True
    calculate_comprehensive_score: bool = True
    exclude_invalid_entries: bool = True
    strict_validation: bool = True
    backup_original: bool = True
    
    def __post_init__(self):
        """验证配置"""
        if not self.input_file:
            raise ValueError("输入文件路径不能为空")
        if not self.output_file:
            raise ValueError("输出文件路径不能为空")
        if self.operation_type not in ["monthly", "final"]:
            raise ValueError("操作类型必须是 'monthly' 或 'final'")
    
    @property
    def is_monthly_operation(self) -> bool:
        """是否为首月评分操作"""
        return self.operation_type == "monthly"
    
    @property
    def is_final_operation(self) -> bool:
        """是否为完结评分操作"""
        return self.operation_type == "final"
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            "input_file": self.input_file,
            "output_file": self.output_file,
            "operation_type": self.operation_type,
            "preserve_hyperlinks": self.preserve_hyperlinks,
            "apply_styles": self.apply_styles,
            "calculate_comprehensive_score": self.calculate_comprehensive_score,
            "exclude_invalid_entries": self.exclude_invalid_entries,
            "strict_validation": self.strict_validation,
            "backup_original": self.backup_original,
        }
    
    @classmethod
    def for_monthly_processing(cls, input_file: str, output_file: str) -> 'ProcessingConfig':
        """创建首月评分处理配置"""
        return cls(
            input_file=input_file,
            output_file=output_file,
            operation_type="monthly"
        )
    
    @classmethod
    def for_final_processing(cls, input_file: str, output_file: str) -> 'ProcessingConfig':
        """创建完结评分处理配置"""
        return cls(
            input_file=input_file,
            output_file=output_file,
            operation_type="final"
        ) 