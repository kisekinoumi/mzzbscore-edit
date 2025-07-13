"""
排名处理服务

提供数据过滤、评分计算和排名处理的服务。
"""

import pandas as pd
import numpy as np
import time
from typing import Dict, List, Optional, Tuple

from app.core.base import BaseService, IDataProcessor
from app.models.data_models import RankingResult, AnimeData
from app.config.constants import (
    ExcelColumns, EXCLUDED_NOTES, PLATFORM_COLUMNS, 
    COMPREHENSIVE_SCORE_WEIGHTS, COMPREHENSIVE_RANKING_COLUMN
)
from app.utils.exceptions import RankingError, ScoreCalculationError, ValidationError
from app.utils.validators import DataValidator
from app.utils.logger import Logger


class RankingService(BaseService, IDataProcessor):
    """排名处理服务"""
    
    def __init__(self, logger: Optional[Logger] = None):
        """
        初始化排名服务
        
        Args:
            logger: 日志记录器
        """
        super().__init__(logger.get_logger("RankingService") if logger else None)
        self._excluded_entries: Optional[pd.DataFrame] = None
        self._ranking_config = {
            "method": "min",  # 排名方法: min, max, average, first, dense
            "ascending": False,  # 分数越高排名越靠前
            "na_option": "keep"  # NaN值的处理方式
        }
    
    def initialize(self) -> bool:
        """
        初始化服务
        
        Returns:
            bool: 初始化是否成功
        """
        try:
            self.logger.info("RankingService初始化开始")
            
            # 验证配置
            if not COMPREHENSIVE_SCORE_WEIGHTS:
                raise RankingError("综合评分权重配置为空")
            
            if not PLATFORM_COLUMNS:
                raise RankingError("平台列配置为空")
            
            # 验证权重总和
            total_weight = sum(COMPREHENSIVE_SCORE_WEIGHTS.values())
            if abs(total_weight - 1.0) > 0.01:  # 允许小的浮点误差
                self.logger.warning(f"综合评分权重总和不为1.0: {total_weight}")
            
            self._set_initialized(True)
            self.logger.info("RankingService初始化成功")
            return True
            
        except Exception as e:
            self.logger.error(f"RankingService初始化失败: {e}")
            return False
    
    def validate_data(self, data: pd.DataFrame) -> bool:
        """
        验证数据格式
        
        Args:
            data: 要验证的数据
            
        Returns:
            bool: 数据是否有效
        """
        try:
            # 基本验证
            DataValidator.validate_dataframe_not_empty(data, "输入数据")
            
            # 验证必需的列
            required_columns = [ExcelColumns.ORIGINAL_NAME, ExcelColumns.NOTES]
            DataValidator.validate_required_columns(data, required_columns)
            
            # 验证数据完整性
            DataValidator.validate_data_integrity(data, ExcelColumns.ORIGINAL_NAME)
            
            self.logger.info("数据验证通过")
            return True
            
        except ValidationError as e:
            self.logger.error(f"数据验证失败: {e}")
            raise
        except Exception as e:
            self.logger.error(f"数据验证时发生错误: {e}")
            raise ValidationError(f"数据验证失败: {e}")
    
    def process_data(self, data: pd.DataFrame) -> pd.DataFrame:
        """
        处理数据（简化版本，主要用于接口实现）
        
        Args:
            data: 输入数据
            
        Returns:
            pd.DataFrame: 处理后的数据
        """
        result = self.process_rankings(data)
        return result.valid_data
    
    def process_rankings(self, data: pd.DataFrame) -> RankingResult:
        """
        执行完整的排名处理流程
        
        Args:
            data: 输入数据
            
        Returns:
            RankingResult: 排名结果
        """
        start_time = time.time()
        errors = []
        warnings = []
        
        try:
            self.logger.info("开始处理排名数据")
            
            # 验证数据
            self.validate_data(data)
            
            # 1. 过滤数据
            try:
                valid_data = self._filter_entries(data.copy())
                if valid_data.empty:
                    warnings.append("过滤后没有有效数据")
                    return RankingResult(
                        valid_data=valid_data,
                        excluded_data=self._excluded_entries if self._excluded_entries is not None else pd.DataFrame(),
                        processing_time=time.time() - start_time,
                        warnings=warnings
                    )
            except Exception as e:
                error_msg = f"数据过滤失败: {e}"
                errors.append(error_msg)
                self.logger.error(error_msg)
                # 继续使用原始数据
                valid_data = data.copy()
                self._excluded_entries = pd.DataFrame()
            
            # 2. 计算综合评分
            try:
                valid_data = self._calculate_comprehensive_score(valid_data)
            except Exception as e:
                error_msg = f"综合评分计算失败: {e}"
                errors.append(error_msg)
                self.logger.error(error_msg)
            
            # 3. 计算综合排名
            try:
                if ExcelColumns.COMPREHENSIVE_SCORE in valid_data.columns:
                    valid_data = self._calculate_ranking(
                        valid_data, 
                        ExcelColumns.COMPREHENSIVE_SCORE, 
                        COMPREHENSIVE_RANKING_COLUMN
                    )
            except Exception as e:
                error_msg = f"综合排名计算失败: {e}"
                errors.append(error_msg)
                self.logger.error(error_msg)
            
            # 4. 计算各平台排名
            platform_errors = []
            for platform, (score_col, rank_col, _) in PLATFORM_COLUMNS.items():
                try:
                    if score_col in valid_data.columns:
                        valid_data = self._calculate_ranking(valid_data, score_col, rank_col)
                    else:
                        warning_msg = f"平台 '{platform}' 的评分列 '{score_col}' 不存在"
                        warnings.append(warning_msg)
                        self.logger.warning(warning_msg)
                        # 确保排名列存在
                        valid_data[rank_col] = pd.NA
                except Exception as e:
                    error_msg = f"平台 '{platform}' 排名计算失败: {e}"
                    platform_errors.append(error_msg)
                    self.logger.error(error_msg)
                    # 确保排名列存在
                    try:
                        valid_data[rank_col] = pd.NA
                    except:
                        pass
            
            if platform_errors:
                errors.extend(platform_errors)
            
            # 5. 为排除的条目添加排名列
            try:
                self._add_ranking_columns_to_excluded(valid_data)
            except Exception as e:
                warning_msg = f"为排除条目添加排名列时发生错误: {e}"
                warnings.append(warning_msg)
                self.logger.warning(warning_msg)
            
            # 创建结果
            processing_time = time.time() - start_time
            result = RankingResult(
                valid_data=valid_data,
                excluded_data=self._excluded_entries if self._excluded_entries is not None else pd.DataFrame(),
                processing_time=processing_time,
                errors=errors,
                warnings=warnings
            )
            
            # 记录统计信息
            summary = result.get_summary()
            self.logger.info(f"排名处理完成: {summary}")
            
            return result
            
        except Exception as e:
            processing_time = time.time() - start_time
            error_msg = f"排名处理过程中发生严重错误: {e}"
            errors.append(error_msg)
            self.logger.error(error_msg, exc_info=True)
            
            # 返回基本结果
            return RankingResult(
                valid_data=data.copy() if data is not None else pd.DataFrame(),
                excluded_data=pd.DataFrame(),
                processing_time=processing_time,
                errors=errors,
                warnings=warnings
            )
    
    def _filter_entries(self, data: pd.DataFrame) -> pd.DataFrame:
        """
        根据Notes列过滤条目
        
        Args:
            data: 输入数据
            
        Returns:
            pd.DataFrame: 过滤后的有效数据
        """
        try:
            self.logger.info("开始过滤条目")
            notes_col = ExcelColumns.NOTES
            
            if notes_col not in data.columns:
                self.logger.warning(f"列 '{notes_col}' 不存在，跳过过滤")
                self._excluded_entries = pd.DataFrame()
                return data
            
            # 过滤逻辑
            excluded_mask = data[notes_col].fillna('').isin(EXCLUDED_NOTES)
            self._excluded_entries = data[excluded_mask].copy()
            filtered_data = data[~excluded_mask].copy()
            
            self.logger.info(f"过滤完成: 有效条目 {len(filtered_data)}, 排除条目 {len(self._excluded_entries)}")
            
            if len(filtered_data) == 0:
                self.logger.warning("过滤后没有有效条目")
            
            return filtered_data
            
        except Exception as e:
            self.logger.error(f"过滤条目时发生错误: {e}")
            # 如果过滤失败，返回原始数据
            self._excluded_entries = pd.DataFrame()
            return data
    
    def _calculate_comprehensive_score(self, data: pd.DataFrame) -> pd.DataFrame:
        """
        计算综合评分
        
        Args:
            data: 输入数据
            
        Returns:
            pd.DataFrame: 包含综合评分的数据
        """
        try:
            self.logger.info("开始计算综合评分")
            
            if data.empty:
                self.logger.warning("输入数据为空，跳过综合评分计算")
                return data
            
            # 创建综合评分列
            data[ExcelColumns.COMPREHENSIVE_SCORE] = pd.NA
            
            successful_calculations = 0
            failed_calculations = 0
            
            for idx, row in data.iterrows():
                try:
                    valid_scores = {}
                    valid_weights = {}
                    
                    # 收集有效的评分和权重
                    for score_col, weight in COMPREHENSIVE_SCORE_WEIGHTS.items():
                        if score_col not in data.columns:
                            continue
                        
                        try:
                            score_value = row[score_col]
                            numeric_score = pd.to_numeric(score_value, errors='raise')
                            
                            if not pd.isna(numeric_score) and numeric_score >= 0:
                                valid_scores[score_col] = numeric_score
                                valid_weights[score_col] = weight
                        except (ValueError, TypeError):
                            continue
                    
                    # 计算加权平均分
                    if valid_scores:
                        total_weight = sum(valid_weights.values())
                        if total_weight > 0:
                            weighted_sum = sum(
                                score * valid_weights[col] 
                                for col, score in valid_scores.items()
                            )
                            comprehensive_score = weighted_sum / total_weight
                            
                            # 验证结果
                            if not pd.isna(comprehensive_score) and comprehensive_score >= 0:
                                data.at[idx, ExcelColumns.COMPREHENSIVE_SCORE] = comprehensive_score
                                successful_calculations += 1
                                
                                # 调试信息
                                used_platforms = list(valid_scores.keys())
                                self.logger.debug(
                                    f"条目 {row.get(ExcelColumns.ORIGINAL_NAME, 'Unknown')} "
                                    f"使用平台: {used_platforms}, 综合评分: {comprehensive_score:.2f}"
                                )
                            else:
                                failed_calculations += 1
                        else:
                            failed_calculations += 1
                    else:
                        # 没有有效评分数据
                        self.logger.debug(
                            f"条目 {row.get(ExcelColumns.ORIGINAL_NAME, 'Unknown')} 没有有效的评分数据"
                        )
                        
                except Exception as e:
                    self.logger.warning(f"处理条目 {idx} 时发生错误: {e}")
                    failed_calculations += 1
            
            self.logger.info(
                f"综合评分计算完成: 成功 {successful_calculations}, 失败 {failed_calculations}"
            )
            return data
            
        except Exception as e:
            self.logger.error(f"计算综合评分时发生错误: {e}")
            raise ScoreCalculationError(f"综合评分计算失败: {e}")
    
    def _calculate_ranking(self, data: pd.DataFrame, score_col: str, rank_col: str) -> pd.DataFrame:
        """
        计算指定列的排名
        
        Args:
            data: 数据
            score_col: 评分列名
            rank_col: 排名列名
            
        Returns:
            pd.DataFrame: 包含排名的数据
        """
        try:
            self.logger.debug(f"开始计算 '{score_col}' 列的排名")
            
            if data.empty:
                self.logger.warning("数据为空，跳过排名计算")
                return data
            
            if score_col not in data.columns:
                self.logger.warning(f"列 '{score_col}' 不存在，跳过排名计算")
                data[rank_col] = pd.NA
                return data
            
            # 转换为数值类型
            scores = pd.to_numeric(data[score_col], errors='coerce')
            
            # 检查有效数值
            valid_scores = scores.dropna()
            if len(valid_scores) == 0:
                self.logger.warning(f"列 '{score_col}' 中没有有效数值，无法计算排名")
                data[rank_col] = pd.NA
                return data
            
            # 计算排名
            ranks = scores.rank(
                method=self._ranking_config["method"],
                ascending=self._ranking_config["ascending"],
                na_option=self._ranking_config["na_option"]
            )
            
            # 转换为整数类型（支持NaN）
            data[rank_col] = ranks.astype('Int64')
            
            # 统计
            ranked_count = data[rank_col].notna().sum()
            self.logger.debug(f"'{score_col}' 列排名完成，共 {ranked_count} 个条目获得排名")
            
            return data
            
        except Exception as e:
            self.logger.error(f"计算排名时发生错误 (列: {score_col}): {e}")
            # 确保排名列存在
            data[rank_col] = pd.NA
            return data
    
    def _add_ranking_columns_to_excluded(self, valid_data: pd.DataFrame):
        """
        为排除的条目添加排名列
        
        Args:
            valid_data: 有效数据（用于获取列结构）
        """
        try:
            if self._excluded_entries is None or self._excluded_entries.empty:
                return
            
            # 添加所有排名列
            rank_cols = [cols[1] for cols in PLATFORM_COLUMNS.values()]
            for rank_col in rank_cols:
                if rank_col in valid_data.columns:
                    self._excluded_entries[rank_col] = pd.NA
            
            # 添加综合评分和综合排名列
            if ExcelColumns.COMPREHENSIVE_SCORE in valid_data.columns:
                self._excluded_entries[ExcelColumns.COMPREHENSIVE_SCORE] = pd.NA
            if COMPREHENSIVE_RANKING_COLUMN in valid_data.columns:
                self._excluded_entries[COMPREHENSIVE_RANKING_COLUMN] = pd.NA
                
        except Exception as e:
            self.logger.warning(f"为排除条目添加排名列时发生错误: {e}")
    
    def get_ranking_statistics(self, result: RankingResult) -> Dict[str, any]:
        """
        获取排名统计信息
        
        Args:
            result: 排名结果
            
        Returns:
            Dict[str, any]: 统计信息
        """
        try:
            stats = {
                "total_entries": result.total_processed,
                "valid_entries": result.total_valid,
                "excluded_entries": result.total_excluded,
                "success_rate": result.success_rate,
                "processing_time": result.processing_time,
                "platform_stats": {}
            }
            
            # 统计各平台的排名情况
            if not result.valid_data.empty:
                for platform, (score_col, rank_col, _) in PLATFORM_COLUMNS.items():
                    if rank_col in result.valid_data.columns:
                        ranked_count = result.valid_data[rank_col].notna().sum()
                        total_count = len(result.valid_data)
                        stats["platform_stats"][platform] = {
                            "ranked_count": ranked_count,
                            "total_count": total_count,
                            "coverage_rate": ranked_count / total_count if total_count > 0 else 0
                        }
            
            # 综合评分统计
            if ExcelColumns.COMPREHENSIVE_SCORE in result.valid_data.columns:
                comp_scores = result.valid_data[ExcelColumns.COMPREHENSIVE_SCORE].dropna()
                if not comp_scores.empty:
                    stats["comprehensive_score_stats"] = {
                        "count": len(comp_scores),
                        "mean": float(comp_scores.mean()),
                        "median": float(comp_scores.median()),
                        "min": float(comp_scores.min()),
                        "max": float(comp_scores.max()),
                        "std": float(comp_scores.std())
                    }
            
            return stats
            
        except Exception as e:
            self.logger.error(f"获取排名统计信息时发生错误: {e}")
            return {"error": str(e)}
    
    def set_ranking_method(self, method: str):
        """
        设置排名方法
        
        Args:
            method: 排名方法 ('min', 'max', 'average', 'first', 'dense')
        """
        valid_methods = ['min', 'max', 'average', 'first', 'dense']
        if method not in valid_methods:
            raise ValueError(f"无效的排名方法: {method}，有效值: {valid_methods}")
        
        self._ranking_config["method"] = method
        self.logger.info(f"排名方法已设置为: {method}")
    
    def get_excluded_entries(self) -> Optional[pd.DataFrame]:
        """
        获取被排除的条目
        
        Returns:
            Optional[pd.DataFrame]: 被排除的条目
        """
        return self._excluded_entries.copy() if self._excluded_entries is not None else None 