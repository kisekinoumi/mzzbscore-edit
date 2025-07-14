import pandas as pd
import openpyxl

print("=== 检查修改后的NaN文本写入逻辑 ===")

# 读取生成的Excel文件
df = pd.read_excel('monthly_anime_scores.xlsx', header=1)

print("1. 平台排名统计信息:")
print("   - Bangumi: 48/48 (100.0%)")
print("   - Anilist: 48/48 (100.0%)")
print("   - MyAnimeList: 48/48 (100.0%)")
print("   - Filmarks: 44/48 (91.7%) <- 4个有效条目没有Filmarks评分")

rank_cols = ['Bangumi_Rank', 'Anilist_Rank', 'Myanimelist_Rank', 'Filmarks_Rank']

print(f"\n2. 检查NaN文本统计:")
for col in rank_cols:
    if col in df.columns:
        nan_text_count = (df[col] == "NaN").sum()
        null_count = df[col].isnull().sum()
        numeric_count = df[col].apply(lambda x: isinstance(x, (int, float)) and not pd.isna(x)).sum()
        total_count = len(df)
        print(f"   {col}:")
        print(f"     - 'NaN'文本: {nan_text_count} 条")
        print(f"     - 空值: {null_count} 条") 
        print(f"     - 数字排名: {numeric_count} 条")
        print(f"     - 总计: {total_count} 条")

print(f"\n3. 重点检查Filmarks_Rank:")
print("   根据统计，应该有4个有效条目没有Filmarks评分，Filmarks_Rank应该为'NaN'文本")

# 筛选出有综合评分但Filmarks_Rank为NaN的条目
has_comprehensive_score = df['综合评分'].notna()
filmarks_rank_is_nan = df['Filmarks_Rank'] == "NaN"

target_entries = df[has_comprehensive_score & filmarks_rank_is_nan]
print(f"   找到 {len(target_entries)} 个有综合评分但Filmarks_Rank为'NaN'的条目:")

if not target_entries.empty:
    display_cols = ['原名', '综合评分', 'Filmarks', 'Filmarks_Rank']
    print(target_entries[display_cols].to_string(index=False))
else:
    print("   没有找到符合条件的条目")

print(f"\n4. 检查有综合评分的条目总数:")
valid_entries = df[df['综合评分'].notna()]
print(f"   有综合评分的条目: {len(valid_entries)} 条")

print(f"\n5. 分别检查各站点在有效条目中的排名情况:")
for col in rank_cols:
    if col in df.columns:
        # 在有综合评分的条目中
        valid_rank_count = valid_entries[col].apply(lambda x: isinstance(x, (int, float)) and not pd.isna(x)).sum()
        valid_nan_count = (valid_entries[col] == "NaN").sum()
        print(f"   {col}: {valid_rank_count}个数字排名 + {valid_nan_count}个'NaN'文本 = {valid_rank_count + valid_nan_count}/48")

print("\n=== 检查完成 ===") 