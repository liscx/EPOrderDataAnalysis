import pandas as pd
import yaml
import os


def load_config():
    """加载YAML配置文件"""
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def run_supplier_analysis():
    # 1. 加载配置
    config = load_config()
    file_path = config['file_config']['output_file']

    if not os.path.exists(file_path):
        print(f"错误：找不到文件 {file_path}，请先运行数据清洗脚本。")
        return

    # 2. 加载数据：全量汇总固定读取“清洗后数据”
    print(f"--- 正在执行：供应商全量汇总分析 ---")
    xls = pd.ExcelFile(file_path)
    # 读取所有现有 Sheet 以便完整保留原有数据
    all_sheets = {name: xls.parse(name) for name in xls.sheet_names}

    if '清洗后数据' not in all_sheets:
        print("错误：未找到 '清洗后数据' 工作表。")
        return

    df = all_sheets['清洗后数据'].copy()

    # 3. 预处理：处理合并单元格产生的空值
    # 使用 ffill 确保每一行商品都能匹配到对应的供应商和订单号
    df['订单号'] = df['订单号'].ffill()
    df['供应商'] = df['供应商'].ffill()
    df['专区名称'] = df['专区名称'].ffill()

    if df.empty:
        print("警告：'清洗后数据' 中无有效记录。")
        return

    # 4. 供应商维度聚合统计
    supplier_report = df.groupby('供应商').agg(
        订单数量=('订单号', 'nunique'),
        订单总额_元=('订单金额（元）', 'sum'),
        商品数量=('数量', 'sum'),
        专区名称=('专区名称', lambda x: "、".join(x.dropna().astype(str).unique()))
    ).reset_index()

    # 5. 排序与格式化
    # 按照订单总额从高到低排列（降序）
    supplier_report = supplier_report.sort_values(by='订单总额_元', ascending=False)

    # 重命名列名
    supplier_report.columns = ['供应商', '订单数量', '订单总额（元）', '商品数量', '专区名称']

    # 插入序号
    supplier_report.insert(0, '序号', range(1, len(supplier_report) + 1))

    # 6. 统一写回 Excel
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        # 先写回所有原始 Sheet (包括清洗数据、区间数据、企业汇总等)
        for name, sheet_df in all_sheets.items():
            # 避开即将生成的供应商汇总表，防止重复
            if name != '供应商汇总表':
                sheet_df.to_excel(writer, sheet_name=name, index=False)

        # 追加/覆盖“供应商汇总表”
        supplier_report.to_excel(writer, sheet_name='供应商汇总表', index=False)

    print(f"全量分析完成！结果已更新至 '{file_path}' 的 '供应商汇总表'。")


if __name__ == "__main__":
    run_supplier_analysis()