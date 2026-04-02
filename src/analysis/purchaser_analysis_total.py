import pandas as pd
import yaml
import os


from ..core import load_config


def run_purchaser_analysis():
    # 1. 加载基础配置
    try:
        config = load_config()
    except Exception as e:
        print(f"错误：加载配置失败: {e}")
        return

    file_path = config['file_config']['output_file']

    if not os.path.exists(file_path):
        print(f"错误：找不到文件 {file_path}，请先运行数据清洗脚本。")
        return

    # 2. 加载数据
    print(f"--- 正在执行：采购主体全量汇总分析 ---")
    with pd.ExcelFile(file_path) as xls:
        sheet_names = xls.sheet_names
        if '清洗后数据' not in sheet_names:
            print("错误：Excel中不存在 '清洗后数据' Sheet。")
            return

        df = pd.read_excel(xls, sheet_name='清洗后数据')

        # 预加载其他 Sheet 用于最后写回
        other_sheets = {}
        for sn in sheet_names:
            if sn != '采购企业汇总表':
                other_sheets[sn] = pd.read_excel(xls, sheet_name=sn)

    if df.empty:
        print("警告：'清洗后数据' 中无有效记录。")
        return

    # --- 核心修复：列名清洗与自动识别 ---
    # 去除列名两端的空格（防止 " 采购企业" 这种情况导致报错）
    df.columns = [str(c).strip() for c in df.columns]

    # 自动判定统计维度：优先寻找“采购企业”，若没有则寻找“采购部门”
    target_col = None
    if '采购企业' in df.columns:
        target_col = '采购企业'
    elif '采购部门' in df.columns:
        target_col = '采购部门'
    else:
        print("错误：原始数据中既没有 '采购企业' 也没有 '采购部门' 列，无法分析。")
        return

    print(f"检测到分析维度列为: [{target_col}]")

    # 3. 预处理：处理合并单元格产生的空值
    df['订单号'] = df['订单号'].ffill()
    df[target_col] = df[target_col].replace(['nan', 'None', ''], pd.NA).ffill()

    # 选配：专区名称如果也存在合并单元格，同样需要填充
    if '专区名称' in df.columns:
        df['专区名称'] = df['专区名称'].ffill()
    else:
        df['专区名称'] = "默认专区"

    df['订单日期'] = pd.to_datetime(df['订单日期'], errors='coerce')
    # 确保金额和数量是数值
    df['订单金额（元）'] = pd.to_numeric(df['订单金额（元）'], errors='coerce').fillna(0)

    # --- 诊断：查看列状态 ---
    print(f"[诊断] 识别到的维度列为: [{target_col}]")
    print(f"[诊断] 该列前5行内容为:\n{df[target_col].head()}")
    print(f"[诊断] 该列空值数量: {df[target_col].isna().sum()}")

    # 4. 汇总概览计算 (增加空数据防御)
    total_ent_count = df[target_col].nunique()
    ent_order_stats = df.groupby(target_col)['订单号'].nunique()
    ent_money_stats = df.groupby(target_col)['订单金额（元）'].sum()

    if not ent_order_stats.empty:
        max_order_val = ent_order_stats.max()
        max_order_ent = ent_order_stats.idxmax()
        max_money_val = ent_money_stats.max()
        max_money_ent = ent_money_stats.idxmax()

        print(f"1. 全量采购主体总数：{total_ent_count} 家")
        print(f"2. 历史最高订单量单位：{max_order_ent} ({max_order_val} 笔)")
        print(f"3. 历史最高交易金额单位：{max_money_ent} ({max_money_val:,.2f} 元)")
    else:
        print(f"⚠️ 预警：识别到的 [{target_col}] 列数据全部为空，无法生成最值统计。")
    print("-" * 50)

    # 5. 构造明细汇总表
    # 注意：这里的 groupby 必须使用我们自动判定的 target_col
    report_df = df.groupby(target_col).agg(
        专区名称=('专区名称', lambda x: " / ".join(map(str, x.dropna().unique()))),
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

    # 统一重命名输出列名（无论原始列叫什么，输出统一叫“采购企业”）
    report_df.columns = ['采购企业', '专区名称', '订单数量', '订单总额（元）', '首次订单日期', '末次订单日期']

    # 过滤掉单位名称为 NaN 或空的结果
    report_df = report_df.dropna(subset=['采购企业'])
    report_df = report_df[report_df['采购企业'] != 'nan']

    # 插入序号
    report_df.insert(0, '序号', range(1, len(report_df) + 1))

    # 6. 统一写回原文件
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            # 先写回备份的原始 Sheet
            for sn, s_df in other_sheets.items():
                s_df.to_excel(writer, sheet_name=sn, index=False)

            # 追加/覆盖“采购企业汇总表”
            report_df.to_excel(writer, sheet_name='采购企业汇总表', index=False)
        print(f"全量分析完成！结果已更新至 '{file_path}' 的 '采购企业汇总表'。")
    except Exception as e:
        print(f"错误：写入 Excel 失败。请确保文件未被打开。详情: {e}")


if __name__ == "__main__":
    run_purchaser_analysis()