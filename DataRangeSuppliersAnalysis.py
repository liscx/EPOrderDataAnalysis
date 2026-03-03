import pandas as pd
import os

# ================= 配置区域 =================
START_DATE = '2026-02-01'
END_DATE = '2026-02-28'
INPUT_FILE = '清洗结果_汇总.xlsx'


# ===========================================

def run_supplier_top10():
    if not os.path.exists(INPUT_FILE):
        print(f"错误：找不到文件 {INPUT_FILE}")
        return

    # 1. 加载所有工作表以备保留
    xls = pd.ExcelFile(INPUT_FILE)
    all_sheets = {name: xls.parse(name) for name in xls.sheet_names}

    df = all_sheets['清洗后数据'].copy()
    df['订单日期'] = pd.to_datetime(df['订单日期'], errors='coerce')

    # 2. 筛选时间区间并处理空值
    mask = (df['订单日期'] >= pd.to_datetime(START_DATE)) & (df['订单日期'] <= pd.to_datetime(END_DATE))
    df_period = df.loc[mask].copy()

    df_period['订单号'] = df_period['订单号'].ffill()
    df_period['供应商'] = df_period['供应商'].ffill()

    # 3. 供应商基础汇总
    sup_summary = df_period.groupby('供应商').agg({
        '订单号': 'nunique',  # 订单数量
        '数量': 'sum',  # 商品数量
        '订单金额（元）': 'sum'  # 订单金额
    }).reset_index()

    # 统一列名并增加预留列
    sup_summary.columns = ['单位名称', '订单数量', '商品数量', '订单金额（元）']
    sup_summary['供应商类型'] = '--'  # 预留列

    # 4. 生成 Top 10 榜单

    # --- 榜单1：订单数量前10，按订单金额降序排列 ---
    top10_qty_pool = sup_summary.nlargest(10, '订单数量')
    top10_by_qty = top10_qty_pool.sort_values(by='订单金额（元）', ascending=False).copy()
    top10_by_qty.insert(0, '序号', range(1, len(top10_by_qty) + 1))
    # 调整列顺序
    cols_qty = ['序号', '单位名称', '供应商类型', '订单数量', '商品数量', '订单金额（元）']
    top10_by_qty = top10_by_qty[cols_qty]

    # --- 榜单2：订单金额前10，按订单数量降序排列 ---
    top10_money_pool = sup_summary.nlargest(10, '订单金额（元）')
    top10_by_money = top10_money_pool.sort_values(by='订单数量', ascending=False).copy()
    top10_by_money.insert(0, '序号', range(1, len(top10_by_money) + 1))
    # 调整列顺序
    cols_money = ['序号', '单位名称', '供应商类型', '订单金额（元）', '订单数量', '商品数量']
    top10_by_money = top10_by_money[cols_money]

    # 5. 写回 Excel (保留原工作表，追加新排行页)
    with pd.ExcelWriter(INPUT_FILE, engine='openpyxl') as writer:
        for name, sheet_df in all_sheets.items():
            sheet_df.to_excel(writer, sheet_name=name, index=False)

        top10_by_qty.to_excel(writer, sheet_name='供应商订单量TOP10', index=False)
        top10_by_money.to_excel(writer, sheet_name='供应商交易额TOP10', index=False)

    print(f"分析完成！已在 {INPUT_FILE} 中生成两个供应商 Top10 排行榜。")
    print(f"时间区间：{START_DATE} 至 {END_DATE}")


if __name__ == "__main__":
    run_supplier_top10()
