import pandas as pd
import os

# ================= 配置区域 =================
INPUT_FILE = '清洗结果_汇总.xlsx'
OUTPUT_FILE = '清洗结果_汇总.xlsx'
# ===========================================

def run_analysis():
    if not os.path.exists(INPUT_FILE):
        print(f"错误：找不到文件 {INPUT_FILE}")
        return

    # 1. 加载数据
    with pd.ExcelFile(INPUT_FILE) as xls:
        df = pd.read_excel(xls, sheet_name='清洗后数据')
        # 读取其他页以便完整写回
        df_removed = pd.read_excel(xls, sheet_name='被清洗的测试数据')
        df_period = pd.read_excel(xls, sheet_name='指定区间数据')

    # 2. 预处理：处理合并单元格产生的空值
    df['订单号'] = df['订单号'].ffill()
    df['采购企业'] = df['采购企业'].ffill()
    df['专区名称'] = df['专区名称'].ffill()
    df['订单日期'] = pd.to_datetime(df['订单日期'], errors='coerce')

    # 3. 汇总数据计算
    # A. 采购企业总数
    total_ent_count = df['采购企业'].nunique()

    # B. 单采购企业累计最高订单数量
    ent_order_stats = df.groupby('采购企业')['订单号'].nunique()
    max_order_val = ent_order_stats.max()
    max_order_ent = ent_order_stats.idxmax()

    # C. 单采购企业订单累计总金额最高
    ent_money_stats = df.groupby('采购企业')['订单金额（元）'].sum()
    max_money_val = ent_money_stats.max()
    max_money_ent = ent_money_stats.idxmax()

    # 控制台打印概览
    print("--- 采购企业分析概览 ---")
    print(f"1. 采购企业总数：{total_ent_count} 家")
    print(f"2. 最高订单量企业：{max_order_ent} ({max_order_val} 笔)")
    print(f"3. 最高总金额企业：{max_money_ent} ({max_money_val:,.2f} 元)")
    print("-" * 50)

    # 4. 构造明细表并进行降序排序
    report_df = df.groupby('采购企业').agg(
        专区名称=('专区名称', lambda x: " / ".join(x.dropna().unique())),
        订单数量=('订单号', 'nunique'),
        订单总额_元=('订单金额（元）', 'sum'),
        首次订单日期=('订单日期', 'min'),
        末次订单日期=('订单日期', 'max')
    ).reset_index()

    # --- 关键优化：按金额降序排列 ---
    report_df = report_df.sort_values(by='订单总额_元', ascending=False)

    # 格式化日期
    report_df['首次订单日期'] = report_df['首次订单日期'].dt.strftime('%Y-%m-%d')
    report_df['末次订单日期'] = report_df['末次订单日期'].dt.strftime('%Y-%m-%d')

    # 重命名列名
    report_df.columns = ['采购企业', '专区名称', '订单数量', '订单总额（元）', '首次订单日期', '末次订单日期']

    # 重新生成序号（按排序后的顺序）
    report_df.insert(0, '序号', range(1, len(report_df) + 1))

    # 5. 保存回原文件，增加/覆盖 Sheet
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='清洗后数据', index=False)
        df_removed.to_excel(writer, sheet_name='被清洗的测试数据', index=False)
        df_period.to_excel(writer, sheet_name='指定区间数据', index=False)
        report_df.to_excel(writer, sheet_name='采购企业汇总表', index=False)

    print(f"分析完成！结果已更新至 '{OUTPUT_FILE}' 的 '采购企业汇总表'，已按金额降序排列。")

if __name__ == "__main__":
    run_analysis()