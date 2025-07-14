import pandas as pd
from app.services.excel_service import ExcelService
from app.services.ranking_service import RankingService
from app.utils.logger import Logger

print("=== 调试排名数据结构 ===")

# 初始化服务
logger = Logger()
excel_service = ExcelService(logger)
ranking_service = RankingService(logger)

excel_service.initialize()
ranking_service.initialize()

# 读取数据
df = excel_service.read_file("mzzb.xlsx")
print(f"1. 读取到 {len(df)} 条记录")

# 处理排名
result = ranking_service.process_rankings(df)
print(f"2. 有效条目: {len(result.valid_data)}")

# 检查有效数据中的列
print(f"3. 有效数据的列: {list(result.valid_data.columns)}")

# 检查Filmarks相关数据
print(f"4. 检查Filmarks相关数据:")
filmarks_cols = [col for col in result.valid_data.columns if 'Filmarks' in col]
print(f"   Filmarks相关列: {filmarks_cols}")

# 检查具体的数据
print(f"5. 前10行Filmarks数据:")
display_cols = ['原名', 'Filmarks', 'Filmarks_Rank'] 
for col in display_cols:
    if col not in result.valid_data.columns:
        print(f"   警告: 列 '{col}' 不存在")

available_cols = [col for col in display_cols if col in result.valid_data.columns]
if available_cols:
    sample = result.valid_data[available_cols].head(10)
    print(sample.to_string())

# 检查哪些条目没有Filmarks_Rank
print(f"\n6. 检查Filmarks_Rank的数据类型和缺失情况:")
if 'Filmarks_Rank' in result.valid_data.columns:
    filmarks_rank_col = result.valid_data['Filmarks_Rank']
    print(f"   数据类型: {filmarks_rank_col.dtype}")
    print(f"   非空值数量: {filmarks_rank_col.count()}")
    print(f"   空值数量: {filmarks_rank_col.isnull().sum()}")
    print(f"   包含NaN字符串的数量: {(filmarks_rank_col == 'NaN').sum()}")
    
    # 查看实际的空值条目
    missing_filmarks_rank = result.valid_data[filmarks_rank_col.isnull()]
    if not missing_filmarks_rank.empty:
        print(f"   没有Filmarks_Rank的条目 ({len(missing_filmarks_rank)}个):")
        print(missing_filmarks_rank[['原名', 'Filmarks']].to_string())
else:
    print("   Filmarks_Rank列不存在!")

print("\n=== 调试完成 ===") 