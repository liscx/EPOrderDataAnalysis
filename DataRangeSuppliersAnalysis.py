# 供应商订单量/交易额 TOP10 分析
import pandas as pd
import yaml
import os


def load_config():
    """加载YAML配置文件"""
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def run_range_supplier_top10():
    # 1. 加载配置
    config = load_config()
    file_path = config['file_config']['output_file']
    start_date = config['analysis_period']['start_date']
    end_date = config['analysis_period']['end_date']

    if not os.path.exists(file_path):
        print(f"错误：找不到文件 {file_path}")
        return

    # 2. 确定数据源
    print(f"--- 正在执行：供应商区间 Top10 分析 ---")
    with pd.ExcelFile(file_path) as xls:
        all_sheets = {name: xls.parse(name) for name in xls.sheet_names}

    if '指定区间数据' in all_sheets:
        print("检测到 '指定区间数据' Sheet，直接读取...")
        df_period = all_sheets['指定区间数据'].copy()
    else:
        print("未检测到预切片数据，正在从 '清洗后数据' 中筛选日期...")
        df = all_sheets['清洗后数据'].copy()
        df['订单日期'] = pd.to_datetime(df['订单日期'], errors='coerce')
        mask = (df['订单日期'] >= pd.to_datetime(start_date)) & \
               (df['订单日期'] <= pd.to_datetime(end_date))
        df_period = df.loc[mask].copy()

    if df_period.empty:
        print(f"警告：在区间 {start_date} 至 {end_date} 内未找到任何供应商数据。")
        return

    # 3. 预处理
    df_period['订单号'] = df_period['订单号'].ffill()
    df_period['供应商'] = df_period['供应商'].astype(str).str.strip().ffill()
    # 确保数值列类型正确
    df_period['订单金额（元）'] = pd.to_numeric(df_period['订单金额（元）'], errors='coerce').fillna(0)
    df_period['数量'] = pd.to_numeric(df_period['数量'], errors='coerce').fillna(0)

    # 4. 供应商基础汇总
    sup_summary = df_period.groupby('供应商').agg({
        '订单号': 'nunique',  # 订单数量
        '数量': 'sum',  # 商品数量
        '订单金额（元）': 'sum'  # 订单金额
    }).reset_index()

    # 统一列名
    sup_summary.columns = ['单位名称', '订单数量', '商品数量', '订单金额（元）']
    sup_summary['供应商类型'] = '--'  # 预留占位

    # 5. 生成 Top 10 榜单 (严格按主维度降序)

    # --- 榜单1：【订单数量】前10，按订单数量 降序排列 ---
    top10_by_qty = sup_summary.sort_values(by='订单数量', ascending=False).head(10).copy()
    top10_by_qty.insert(0, '序号', range(1, len(top10_by_qty) + 1))
    cols_qty = ['序号', '单位名称', '供应商类型', '订单数量', '商品数量', '订单金额（元）']
    top10_by_qty = top10_by_qty[cols_qty]

    # --- 榜单2：【订单金额】前10，按订单金额 降序排列 ---
    top10_by_money = sup_summary.sort_values(by='订单金额（元）', ascending=False).head(10).copy()
    top10_by_money.insert(0, '序号', range(1, len(top10_by_money) + 1))
    cols_money = ['序号', '单位名称', '供应商类型', '订单金额（元）', '订单数量', '商品数量']
    top10_by_money = top10_by_money[cols_money]

    # 6. 写回 Excel
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for name, sheet_df in all_sheets.items():
                if name not in ['供应商订单量TOP10', '供应商交易额TOP10']:
                    sheet_df.to_excel(writer, sheet_name=name, index=False)

            top10_by_qty.to_excel(writer, sheet_name='供应商订单量TOP10', index=False)
            top10_by_money.to_excel(writer, sheet_name='供应商交易额TOP10', index=False)

        print(f"分析完成！已在 {file_path} 中更新供应商 Top10 排行榜。")
        print(f"统计时间区间：{start_date} 至 {end_date}")

    except Exception as e:
        print(f"错误：写入 Excel 失败。请确保文件未被打开。详情: {e}")


if __name__ == "__main__":
    run_range_supplier_top10()