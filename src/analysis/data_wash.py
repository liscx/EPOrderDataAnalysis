import pandas as pd
import os
from ..core import load_config, TEST_KEYWORDS

def do_wash():
    # 1. 加载配置
    config = load_config()
    input_path = config['file_config']['input_file']
    output_path = config['file_config']['output_file']
    start_date = pd.to_datetime(config['analysis_period']['start_date'])
    # 处理 end_date，确保包含结束当天的全部数据 (补全到 23:59:59)
    end_date = pd.to_datetime(config['analysis_period']['end_date']) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

    if not os.path.exists(input_path):
        print(f"错误：找不到原始文件 {input_path}")
        return

    print(f"开始执行数据清洗与区间切片...")

    # 2. 读取原始数据
    df = pd.read_excel(input_path)

    # 3. 预处理：处理合并单元格逻辑 (ffill)
    # 创建辅助列用于锁定整个订单，防止清洗或筛选时把订单里的明细行拆散
    df['辅助订单号'] = df['订单号'].ffill()
    df['辅助日期'] = df['订单日期'].ffill()
    df['辅助日期'] = pd.to_datetime(df['辅助日期'], errors='coerce')

    # 4. 识别测试数据 (清洗逻辑) - 更加严谨的过滤逻辑
    # 关键词已统一移动到 src/core/constants.py
    
    # 针对整个 DataFrame 查找包含关键词的行，避免对数值列进行 join 操作导致的崩溃
    # 逻辑：找出所有字符串类型的列，在这些列中搜索关键词
    str_cols = df.select_dtypes(include=['object']).columns
    mask_hit_test = df[str_cols].apply(
        lambda row: any(row.astype(str).str.contains('|'.join(TEST_KEYWORDS), na=False)), 
        axis=1
    )
    
    # 获取需要剔除的订单号
    test_order_ids = df.loc[mask_hit_test, '辅助订单号'].unique()

    # 拆分：清洗后的全量数据 vs 被移除的测试数据
    df_removed = df[df['辅助订单号'].isin(test_order_ids)].copy()
    df_cleaned_all = df[~df['辅助订单号'].isin(test_order_ids)].copy()

    # 5. 生成“指定区间数据” (切片逻辑)
    # 基于清洗后的数据，利用辅助日期进行区间筛选
    mask_period = (df_cleaned_all['辅助日期'] >= start_date) & \
                  (df_cleaned_all['辅助日期'] <= end_date)
    df_period = df_cleaned_all[mask_period].copy()

    # 6. 清理辅助列，恢复结构
    for temp_df in [df_removed, df_cleaned_all, df_period]:
        temp_df.drop(columns=['辅助订单号', '辅助日期'], inplace=True, errors='ignore')

    # 7. 写入 Excel (作为后续所有脚本的数据源)
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_cleaned_all.to_excel(writer, sheet_name='清洗后数据', index=False)
        df_removed.to_excel(writer, sheet_name='被清洗的测试数据', index=False)
        df_period.to_excel(writer, sheet_name='指定区间数据', index=False)

    print(f"清洗与切片完成！")
    print(f"-> [全量清洗数据]: {len(df_cleaned_all)} 行")
    print(f"-> [指定区间数据]: {len(df_period)} 行 (区间: {config['analysis_period']['start_date']} 至 {config['analysis_period']['end_date']})")
    print(f"-> [拦截测试数据]: {len(df_removed)} 行")
    print(f"结果已保存至: {output_path}")

if __name__ == "__main__":
    do_wash()