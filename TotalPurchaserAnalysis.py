import pandas as pd
import yaml
import os


def load_config():
    """加载YAML配置文件"""
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def run_purchaser_analysis():
    # 1. 加载基础配置
    config = load_config()
    file_path = config['file_config']['output_file']

    if not os.path.exists(file_path):
        print(f"错误：找不到文件 {file_path}，请先运行数据清洗脚本。")
        return

    # 2. 加载数据：全量汇总固定读取“清洗后数据”
    print(f"--- 正在执行：采购企业全量汇总分析 ---")
    with pd.ExcelFile(file_path) as xls:
        # 获取所有 Sheet 名称，以便完整保留数据
        sheet_names = xls.sheet_names

        # 核心分析源：清洗后数据
        if '清洗后数据' not in sheet_names:
            print("错误：Excel中不存在 '清洗后数据' Sheet。")
            return

        df = pd.read_excel(xls, sheet_name='清洗后数据')

        # 预加载其他 Sheet 用于最后写回，防止丢失
        other_sheets = {}
        for sn in sheet_names:
            if sn != '采购企业汇总表':  # 汇总表会被重新生成，不用备份
                other_sheets[sn] = pd.read_excel(xls, sheet_name=sn)

    # 3. 预处理：处理合并单元格产生的空值
    # 使用 ffill 确保每一行商品都能匹配到对应的采购企业和订单号
    df['订单号'] = df['订单号'].ffill()
    df['采购企业'] = df['采购企业'].ffill()
    df['专区名称'] = df['专区名称'].ffill()
    df['订单日期'] = pd.to_datetime(df['订单日期'], errors='coerce')

    if df.empty:
        print("警告：'清洗后数据' 中无有效记录。")
        return

    # 4. 汇总概览计算 (控制台输出)
    total_ent_count = df['采购企业'].nunique()

    # 累计最高订单量企业
    ent_order_stats = df.groupby('采购企业')['订单号'].nunique()
    max_order_val = ent_order_stats.max()
    max_order_ent = ent_order_stats.idxmax()

    # 累计最高总金额企业
    ent_money_stats = df.groupby('采购企业')['订单金额（元）'].sum()
    max_money_val = ent_money_stats.max()
    max_money_ent = ent_money_stats.idxmax()

    print(f"1. 全量采购企业总数：{total_ent_count} 家")
    print(f"2. 历史最高订单量企业：{max_order_ent} ({max_order_val} 笔)")
    print(f"3. 历史最高交易金额企业：{max_money_ent} ({max_money_val:,.2f} 元)")
    print("-" * 50)

    # 5. 构造明细汇总表
    report_df = df.groupby('采购企业').agg(
        专区名称=('专区名称', lambda x: " / ".join(x.dropna().unique())),
        订单数量=('订单号', 'nunique'),
        订单总额_元=('订单金额（元）', 'sum'),
        首次订单日期=('订单日期', 'min'),
        末次订单日期=('订单日期', 'max')
    ).reset_index()

    # 按金额降序排列
    report_df = report_df.sort_values(by='订单总额_元', ascending=False)

    # 格式化日期显示
    report_df['首次订单日期'] = report_df['首次订单日期'].dt.strftime('%Y-%m-%d')
    report_df['末次订单日期'] = report_df['末次订单日期'].dt.strftime('%Y-%m-%d')

    # 重命名列名
    report_df.columns = ['采购企业', '专区名称', '订单数量', '订单总额（元）', '首次订单日期', '末次订单日期']

    # 插入序号
    report_df.insert(0, '序号', range(1, len(report_df) + 1))

    # 6. 统一写回原文件
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        # 先写回备份的原始 Sheet (清洗后数据、指定区间数据、被清洗数据等)
        for sn, s_df in other_sheets.items():
            s_df.to_excel(writer, sheet_name=sn, index=False)

        # 追加/覆盖“采购企业汇总表”
        report_df.to_excel(writer, sheet_name='采购企业汇总表', index=False)

    print(f"全量分析完成！结果已更新至 '{file_path}' 的 '采购企业汇总表'。")


if __name__ == "__main__":
    run_purchaser_analysis()