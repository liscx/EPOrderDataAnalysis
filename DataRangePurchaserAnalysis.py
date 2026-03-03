import pandas as pd
import yaml
import os


def load_config():
    """加载YAML配置文件"""
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def run_range_purchaser_top10():
    # 1. 加载配置
    config = load_config()
    file_path = config['file_config']['output_file']
    start_date = config['analysis_period']['start_date']
    end_date = config['analysis_period']['end_date']

    if not os.path.exists(file_path):
        print(f"错误：找不到文件 {file_path}")
        return

    # 2. 确定数据源与筛选逻辑
    print(f"--- 正在执行：采购企业区间 Top10 分析 ---")
    with pd.ExcelFile(file_path) as xls:
        all_sheets = {name: xls.parse(name) for name in xls.sheet_names}

        if '指定区间数据' in all_sheets:
            print("检测到 '指定区间数据' Sheet，直接读取...")
            df_period = all_sheets['指定区间数据'].copy()
        else:
            print("未检测到预切片数据，正在从 '清洗后数据' 中手动筛选日期...")
            df = all_sheets['清洗后数据'].copy()
            df['订单日期'] = pd.to_datetime(df['订单日期'], errors='coerce')
            mask = (df['订单日期'] >= pd.to_datetime(start_date)) & \
                   (df['订单日期'] <= pd.to_datetime(end_date))
            df_period = df.loc[mask].copy()

    if df_period.empty:
        print(f"警告：在区间 {start_date} 至 {end_date} 内未找到任何数据。")
        return

    # 3. 预处理：处理合并单元格空值
    df_period['订单号'] = df_period['订单号'].ffill()
    df_period['采购企业'] = df_period['采购企业'].ffill()
    df_period['专区名称'] = df_period['专区名称'].ffill()

    # 增加“省市”列：提取专区名称的前两个字
    df_period['省市'] = df_period['专区名称'].astype(str).str[:2]

    # 4. 基础汇总（按企业聚合）
    ent_summary = df_period.groupby('采购企业').agg({
        '专区名称': lambda x: " / ".join(x.dropna().unique()),
        '省市': 'first',
        '订单号': 'nunique',  # 订单数量
        '数量': 'sum',  # 商品数量
        '订单金额（元）': 'sum'  # 订单金额
    }).reset_index()

    ent_summary.columns = ['采购企业', '专区名称', '省市', '订单数量', '商品数量', '订单金额（元）']

    # 5. 生成 Top 10 榜单

    # --- 榜单1：按订单数量降序 ---
    top10_by_count = ent_summary.sort_values(by='订单数量', ascending=False).head(10).copy()
    top10_by_count.insert(0, '序号', range(1, len(top10_by_count) + 1))
    cols_count = ['序号', '采购企业', '专区名称', '省市', '订单数量', '商品数量', '订单金额（元）']
    top10_by_count = top10_by_count[cols_count]

    # --- 榜单2：按订单金额降序 ---
    top10_by_money = ent_summary.sort_values(by='订单金额（元）', ascending=False).head(10).copy()
    top10_by_money.insert(0, '序号', range(1, len(top10_by_money) + 1))
    cols_money = ['序号', '采购企业', '专区名称', '省市', '订单金额（元）', '订单数量', '商品数量']
    top10_by_money = top10_by_money[cols_money]

    # 6. 写回 Excel (保留原工作表，追加新排行页)
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for name, sheet_df in all_sheets.items():
            # 过滤掉即将重新生成的 Top10 Sheet
            if name not in ['采购企业订单量TOP10', '采购企业交易额TOP10']:
                sheet_df.to_excel(writer, sheet_name=name, index=False)

        top10_by_count.to_excel(writer, sheet_name='采购企业订单量TOP10', index=False)
        top10_by_money.to_excel(writer, sheet_name='采购企业交易额TOP10', index=False)

    print(f"分析完成！已在 {file_path} 中更新 Top10 排行榜。")
    print(f"统计区间: {start_date} 至 {end_date}")


if __name__ == "__main__":
    run_range_purchaser_top10()