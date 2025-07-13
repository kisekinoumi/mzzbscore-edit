"""
应用程序常量配置

集中管理所有常量定义。
"""

from typing import Dict, List, Tuple


class ExcelColumns:
    """Excel列名配置类，集中定义所有列名"""
    
    # 基本信息列
    ORIGINAL_NAME = "原名"
    TRANSLATED_NAME = "译名"
    
    # 评分列
    BANGUMI_SCORE = "Bangumi"
    BANGUMI_TOTAL = "Bangumi_total"
    BANGUMI_RANK = "Bangumi_Rank"
    
    ANILIST_SCORE = "Anilist"
    ANILIST_TOTAL = "Anilist_total"
    ANILIST_RANK = "Anilist_Rank"
    
    MYANIMELIST_SCORE = "MyAnimelist"
    MYANIMELIST_TOTAL = "MyAnimelist_total"
    MYANIMELIST_RANK = "Myanimelist_Rank"
    
    FILMARKS_SCORE = "Filmarks"
    FILMARKS_TOTAL = "Filmarks_total"
    FILMARKS_RANK = "Filmarks_Rank"
    
    # 综合信息列
    COMPREHENSIVE_SCORE = "综合评分"
    RANKING = "排名"
    
    # 备注列
    NOTES = "Notes"
    
    # X评分列
    X_SCORE = "X"
    X_FAN = "X_fan"
    
    # URL链接列
    BANGUMI_URL = "Bangumi_url"
    ANILIST_URL = "Anilist_url"
    MYANILIST_URL = "Myanilist_url"
    FILMARKS_URL = "Filmarks_url"


# 需要从排序中排除的条目备注
EXCLUDED_NOTES: List[str] = ["*时长不足", "*数据不足"]

# 综合评分权重配置
COMPREHENSIVE_SCORE_WEIGHTS: Dict[str, float] = {
    ExcelColumns.BANGUMI_SCORE: 0.5,
    ExcelColumns.ANILIST_SCORE: 0.2,
    ExcelColumns.MYANIMELIST_SCORE: 0.1,
    ExcelColumns.FILMARKS_SCORE: 0.2
}

# 需要进行排名的平台列
PLATFORM_COLUMNS: Dict[str, Tuple[str, str, str]] = {
    "Bangumi": (ExcelColumns.BANGUMI_SCORE, ExcelColumns.BANGUMI_RANK, ExcelColumns.BANGUMI_TOTAL),
    "Anilist": (ExcelColumns.ANILIST_SCORE, ExcelColumns.ANILIST_RANK, ExcelColumns.ANILIST_TOTAL),
    "MyAnimeList": (ExcelColumns.MYANIMELIST_SCORE, ExcelColumns.MYANIMELIST_RANK, ExcelColumns.MYANIMELIST_TOTAL),
    "Filmarks": (ExcelColumns.FILMARKS_SCORE, ExcelColumns.FILMARKS_RANK, ExcelColumns.FILMARKS_TOTAL)
}

# 综合评分使用现有的"排名"列，不插入新列
COMPREHENSIVE_RANKING_COLUMN: str = ExcelColumns.RANKING

# Excel样式配置
EXCEL_STYLE_CONFIG: Dict[str, Dict[str, str]] = {
    "基本信息": {
        "columns": [ExcelColumns.ORIGINAL_NAME, ExcelColumns.TRANSLATED_NAME],
        "color": "E8F4FD"  # 浅蓝色
    },
    "Bangumi": {
        "columns": [ExcelColumns.BANGUMI_SCORE, ExcelColumns.BANGUMI_TOTAL, ExcelColumns.BANGUMI_RANK],
        "color": "E8F8E8"  # 浅绿色
    },
    "Anilist": {
        "columns": [ExcelColumns.ANILIST_SCORE, ExcelColumns.ANILIST_TOTAL, ExcelColumns.ANILIST_RANK],
        "color": "FFF2E8"  # 浅橙色
    },
    "MyAnimelist": {
        "columns": [ExcelColumns.MYANIMELIST_SCORE, ExcelColumns.MYANIMELIST_TOTAL, ExcelColumns.MYANIMELIST_RANK],
        "color": "F8E8F8"  # 浅紫色
    },
    "Filmarks": {
        "columns": [ExcelColumns.FILMARKS_SCORE, ExcelColumns.FILMARKS_TOTAL, ExcelColumns.FILMARKS_RANK],
        "color": "F8F8E8"  # 浅黄色
    },
    "综合评分": {
        "columns": [ExcelColumns.COMPREHENSIVE_SCORE, ExcelColumns.RANKING],
        "color": "E8E8F8"  # 浅紫蓝色
    },
    "X评分": {
        "columns": [ExcelColumns.X_SCORE, ExcelColumns.X_FAN],
        "color": "F0F0F0"  # 浅灰色
    },
    "链接": {
        "columns": [ExcelColumns.BANGUMI_URL, ExcelColumns.ANILIST_URL, ExcelColumns.MYANILIST_URL, ExcelColumns.FILMARKS_URL],
        "color": "E8F8F0"  # 浅青色
    },
    "备注": {
        "columns": [ExcelColumns.NOTES],
        "color": "FFF8E8"  # 浅米色
    }
}

# 文件配置
DEFAULT_INPUT_FILE: str = "mzzb.xlsx"
DEFAULT_OUTPUT_FILE_MONTHLY: str = "monthly_anime_scores.xlsx"
DEFAULT_OUTPUT_FILE_FINAL: str = "final_anime_scores.xlsx"

# 应用程序配置
APP_NAME: str = "Excel动漫评分排名系统"
APP_VERSION: str = "2.0.0"
SUPPORTED_FILE_EXTENSIONS: List[str] = [".xlsx", ".xls"] 