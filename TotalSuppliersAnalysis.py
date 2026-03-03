import pandas as pd
import os

# ================= 配置区域 =================
INPUT_FILE = '清洗结果_汇总.xlsx'
OUTPUT_FILE = '清洗结果_汇总.xlsx'


# ===========================================

def run_supplier_analysis():
    if not os.path.exists(INPUT_FILE):
        print(f"错误：找不到文件 {INPUT_FILE}")
        return

    # 1. 加载数据
    xls = pd.ExcelFile(INPUT_FILE)
    # 读取所有现有 Sheet 以便保留原有数据
    all_sheets = {name: xls.parse(name) for name in xls.sheet_names}

    if '清洗后数据' not in all_sheets:
        print("错误：未找到'清洗后数据'工作表。")
        return

    df = all_sheets['清洗后数据'].copy()

    # 2. 预处理：处理合并单元格产生的空值
    df['订单号'] = df['订单号'].ffill()
    df['供应商'] = df['供应商'].ffill()
    df['专区名称'] = df['专区名称'].ffill()

    # 3. 供应商维度聚合统计
    supplier_report = df.groupby('供应商').agg(
        订单数量=('订单号', 'nunique'),
        订单总额_元=('订单金额（元）', 'sum'),
        商品数量=('数量', 'sum'),
        专区名称=('专区名称', lambda x: "、".join(x.dropna().unique()))  # 使用顿号连接
    ).reset_index()

    # 4. 排序与格式化
    # 按照订单总额从高到低排列
    supplier_report = supplier_report.sort_values(by='订单总额_元', ascending=False)

    # 重命名列名
    supplier_report.columns = ['供应商', '订单数量', '订单总额（元）', '商品数量', '专区名称']

    # 插入序号
    supplier_report.insert(0, '序号', range(1, len(supplier_report) + 1))

    # 5. 写回 Excel
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        for name, sheet_df in all_sheets.items():
            sheet_df.to_excel(writer, sheet_name=name, index=False)

        supplier_report.to_excel(writer, sheet_name='供应商汇总表', index=False)

    print(f"分析完成！结果已更新至 '{OUTPUT_FILE}' 的 '供应商汇总表'。")


if __name__ == "__main__":
    run_supplier_analysis()