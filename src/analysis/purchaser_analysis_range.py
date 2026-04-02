import pandas as pd
import yaml
import os


from ..core import load_config


def run_range_purchaser_top10():
    # 1. 加载配置
    try:
        config = load_config()
    except Exception as e:
        print(f"错误：加载配置失败: {e}")
        return

    file_path = config['file_config']['output_file']
    start_date = config['analysis_period']['start_date']
    end_date = pd.to_datetime(config['analysis_period']['end_date']) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

    if not os.path.exists(file_path):
        print(f"错误：找不到文件 {file_path}")
        return

    # 2. 确定数据源
    print(f"--- 正在执行：采购主体区间 Top10 分析 ---")
    with pd.ExcelFile(file_path) as xls:
        all_sheets = {name: xls.parse(name) for name in xls.sheet_names}

        if '指定区间数据' in all_sheets:
            df_period = all_sheets['指定区间数据'].copy()
        else:
            if '清洗后数据' not in all_sheets:
                print("错误：Excel中缺少数据源。")
                return
            df = all_sheets['清洗后数据'].copy()
            df['订单日期'] = pd.to_datetime(df['订单日期'], errors='coerce')
            mask = (df['订单日期'] >= pd.to_datetime(start_date)) & \
                   (df['订单日期'] <= end_date)
            df_period = df.loc[mask].copy()

    if df_period.empty:
        print(f"警告：在区间 {start_date} 至 {end_date} 内未找到任何数据。")
        return

    # --- 核心修复：列名清洗与自动识别 ---
    df_period.columns = [str(c).strip() for c in df_period.columns]

    # 自动识别采购主体列
    if '采购企业' in df_period.columns:
        target_col = '采购企业'
    elif '采购部门' in df_period.columns:
        target_col = '采购部门'
    else:
        print("错误：未找到 '采购企业' 或 '采购部门' 列。")
        return

    # 3. 预处理与向上填充
    df_period['订单号'] = df_period['订单号'].ffill()
    df_period[target_col] = df_period[target_col].replace(['nan', 'None', ''], pd.NA).ffill()

    if '专区名称' in df_period.columns:
        df_period['专区名称'] = df_period['专区名称'].ffill()
        df_period['省市'] = df_period['专区名称'].astype(str).str[:2]
    else:
        df_period['专区名称'] = "默认专区"
        df_period['省市'] = "--"

    # 4. 金额重算逻辑 (确保明细行金额不丢失)
    df_period['单价（元）'] = pd.to_numeric(df_period['单价（元）'], errors='coerce').fillna(0)
    df_period['数量'] = pd.to_numeric(df_period['数量'], errors='coerce').fillna(0)
    df_period['计算金额'] = df_period['单价（元）'] * df_period['数量']

    # 5. 基础汇总
    ent_summary = df_period.groupby(target_col).agg({
        '专区名称': lambda x: " / ".join(map(str, x.dropna().unique())),
        '省市': 'first',
        '订单号': 'nunique',  # 订单数量（按单号去重）
        '数量': 'sum',  # 商品数量总和
        '计算金额': 'sum'  # 交易金额总和
    }).reset_index()

    # 统一列名输出（保持输出 Sheet 的列名始终为“采购企业”）
    ent_summary.columns = ['采购企业', '专区名称', '省市', '订单数量', '商品数量', '订单金额（元）']

    # 6. 生成 Top 10 榜单
    # 榜单1：按订单数量降序
    top10_by_count = ent_summary.sort_values(by='订单数量', ascending=False).head(10).copy()
    top10_by_count.insert(0, '序号', range(1, len(top10_by_count) + 1))

    # 榜单2：按订单金额降序
    top10_by_money = ent_summary.sort_values(by='订单金额（元）', ascending=False).head(10).copy()
    top10_by_money.insert(0, '序号', range(1, len(top10_by_money) + 1))

    # 7. 写回 Excel
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            # 保留原有的 Sheet
            for name, sheet_df in all_sheets.items():
                if name not in ['采购企业订单量TOP10', '采购企业交易额TOP10']:
                    sheet_df.to_excel(writer, sheet_name=name, index=False)

            # 写入新的排行
            top10_by_count[['序号', '采购企业', '专区名称', '省市', '订单数量', '商品数量', '订单金额（元）']].to_excel(
                writer, sheet_name='采购企业订单量TOP10', index=False)
            top10_by_money[['序号', '采购企业', '专区名称', '省市', '订单金额（元）', '订单数量', '商品数量']].to_excel(
                writer, sheet_name='采购企业交易额TOP10', index=False)

        print(f"分析完成！已在 {file_path} 中更新 Top10 排行榜。")
    except Exception as e:
        print(f"写入失败: {e}")


if __name__ == "__main__":
    run_range_purchaser_top10()