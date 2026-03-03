# step 1
import pandas as pd

# 1. 读取原始 Excel 文件
file_path = '阳光优采交易订单清洗数据.xlsx'
df = pd.read_excel(file_path)

# 2. 预处理：创建辅助列用于识别“合并单元格”关联行
df['辅助订单号'] = df['订单号'].ffill()

# 3. 识别包含关键字的订单
test_keywords = ['测试', '国泰测试','系统管理部']

# 检查整行是否包含关键字
def has_test_keyword(row):
    row_str = "".join(row.astype(str))
    return any(kw in row_str for kw in test_keywords)

# 标记命中行
mask_hit = df.apply(has_test_keyword, axis=1)

# 提取这些订单的完整订单号列表
test_order_ids = df.loc[mask_hit, '辅助订单号'].unique()

# 4. 数据拆分
# Sheet2: 被清洗掉的数据 (命中订单号的所有行)
df_removed = df[df['辅助订单号'].isin(test_order_ids)].copy()

# Sheet1: 清洗后的正常数据
df_cleaned = df[~df['辅助订单号'].isin(test_order_ids)].copy()

# 移除辅助列，恢复原始表头结构
df_removed = df_removed.drop(columns=['辅助订单号'])
df_cleaned = df_cleaned.drop(columns=['辅助订单号'])

# 5. 保存到同一个 Excel 的不同 Sheet 中
output_file = '清洗结果_汇总.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df_cleaned.to_excel(writer, sheet_name='清洗后数据', index=False)
    df_removed.to_excel(writer, sheet_name='被清洗的测试数据', index=False)

# 6. 验证逻辑输出
print(f"处理完成！")
print(f"-> 正常数据已保存至 '清洗后数据' (共 {len(df_cleaned)} 行)")
print(f"-> 测试数据已移至 '被清洗的测试数据' (共 {len(df_removed)} 行)")