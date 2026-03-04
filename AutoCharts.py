import pandas as pd
import yaml
import os
from pyecharts import options as opts
from pyecharts.charts import Pie, Line, Bar, Page
from pyecharts.globals import CurrentConfig

# 设置资源路径，确保图表库正常加载
CurrentConfig.ONLINE_HOST = "https://assets.pyecharts.org/assets/"


def load_config():
    """直接读取本地 config.yaml"""
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def get_base_data(file_path):
    """读取并清洗数据，处理合并单元格"""
    xls = pd.ExcelFile(file_path)
    d_interval = xls.parse("汇总_时间_区间")
    d_full_time = xls.parse("汇总_时间_全量")
    d_full_zone = xls.parse("汇总_专区_全量")
    d_raw = xls.parse("清洗后数据")

    # 填充合并单元格缺失数据
    df_raw_fixed = d_raw.copy()
    for col in ['订单日期', '专区名称', '商品名称']:
        if col in df_raw_fixed.columns:
            df_raw_fixed[col] = df_raw_fixed[col].ffill()

    df_raw_fixed['订单日期'] = pd.to_datetime(df_raw_fixed['订单日期'], errors='coerce')

    return d_interval, d_full_time, d_full_zone, df_raw_fixed


def get_toolbox_opts():
    """通用工具箱设置"""
    return opts.ToolboxOpts(
        is_show=True,
        feature={
            "saveAsImage": {"show": True, "title": "下载图片"},
            "dataView": {"show": True, "title": "数据视图"},
            "restore": {"show": True, "title": "还原"},
        }
    )


def create_pie_component(df, title, attr_col, val_col, is_total_sheet=False):
    """饼图组件"""
    df.columns = [c.strip() for c in df.columns]
    if is_total_sheet:
        mask = (df[attr_col].astype(str).str.contains('小计', na=False)) & \
               (~df[attr_col].astype(str).str.contains('总计|合计|---', na=False))
    else:
        mask = ~df[attr_col].astype(str).str.contains('小计|---', na=False)

    plot_df = df[mask].copy()
    plot_df['Label'] = plot_df[attr_col].astype(str).str.replace(' 小计', '', regex=False)
    data = [(str(row['Label']), round(float(row[val_col]), 2)) for _, row in plot_df.iterrows() if row[val_col] > 0]

    return (
        Pie(init_opts=opts.InitOpts(width="1000px", height="600px", bg_color="white"))
        .add("", data, radius=["35%", "65%"])
        .set_global_opts(
            title_opts=opts.TitleOpts(title=title, pos_left="center"),
            legend_opts=opts.LegendOpts(is_show=False),
            toolbox_opts=get_toolbox_opts()
        )
        .set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {c}元\n({d}%)"))
        .set_colors(["#409EFF", "#67C23A", "#E6A23C", "#F56C6C", "#909399"])
    )


def create_line_component(df, title, x_col=None, y_col=None, is_item_nunique=False):
    """折线图组件 - 已去掉网格线"""
    df_temp = df.copy()

    if not is_item_nunique:
        df_temp.columns = [c.strip() for c in df_temp.columns]
        tdf = df_temp[df_temp[x_col].astype(str).str.contains('小计', na=False)].copy()
        tdf['DT'] = pd.to_datetime(tdf[x_col].astype(str).str.replace(' 小计', ''), format='%Y年%m月', errors='coerce')
        tdf = tdf.dropna(subset=['DT']).set_index('DT').sort_index()
        x_data = [d.strftime('%Y-%m') for d in tdf.index]
        y_data = tdf[y_col].astype(int).tolist()
    else:
        df_temp['Month'] = df_temp['订单日期'].dt.to_period('M').dt.to_timestamp()
        ts = df_temp.groupby('Month')['商品名称'].nunique().sort_index()
        x_data = [d.strftime('%Y-%m') for d in ts.index]
        y_data = ts.astype(int).tolist()

    ml_data = [[{"coord": [x_data[i], y_data[i]]}, {"coord": [x_data[i], 0]}] for i in range(len(x_data))]

    return (
        Line(init_opts=opts.InitOpts(width="1000px", height="500px", bg_color="white"))
        .add_xaxis(x_data)
        .add_yaxis(
            series_name=title, y_axis=y_data, is_smooth=False, symbol="circle", symbol_size=8,
            itemstyle_opts=opts.ItemStyleOpts(color="#409EFF"),
            label_opts=opts.LabelOpts(is_show=True, position="top"),
            markline_opts=opts.MarkLineOpts(
                data=ml_data, symbol=["none", "none"],
                linestyle_opts=opts.LineStyleOpts(type_="dashed", color="#DCDFE6")
            )
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(title=title, pos_left="center"),
            legend_opts=opts.LegendOpts(is_show=False),
            toolbox_opts=get_toolbox_opts(),
            # 去掉网格线
            xaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=False)),
            yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=False)),
        )
    )


def create_bar_component(df, title, x_col, y_col):
    """柱状图组件 - 已去掉网格线"""
    df.columns = [c.strip() for c in df.columns]
    tdf = df[df[x_col].astype(str).str.contains('小计', na=False)].copy()
    tdf['DT'] = pd.to_datetime(tdf[x_col].astype(str).str.replace(' 小计', ''), format='%Y年%m月', errors='coerce')
    tdf = tdf.dropna(subset=['DT']).set_index('DT').sort_index()

    x_data = [d.strftime('%Y-%m') for d in tdf.index]
    y_values = [round(float(v), 2) for v in tdf[y_col]]
    y_axis_max = int(max(y_values) * 1.2) if y_values and max(y_values) > 0 else None

    return (
        Bar(init_opts=opts.InitOpts(width="1000px", height="500px", bg_color="white"))
        .add_xaxis(x_data)
        .add_yaxis(
            "金额", y_values, color="#409EFF", category_gap="45%",
            label_opts=opts.LabelOpts(is_show=True, position="top")
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(title=title, pos_left="center"),
            legend_opts=opts.LegendOpts(is_show=False),
            toolbox_opts=get_toolbox_opts(),
            # 去掉网格线
            xaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=False)),
            yaxis_opts=opts.AxisOpts(max_=y_axis_max, splitline_opts=opts.SplitLineOpts(is_show=False)),
        )
    )


def run_dashboard_output():
    """核心逻辑：读取配置并生成看板"""
    # 1. 直接内部读取配置
    config = load_config()
    excel_path = config['file_config']['output_file']
    output_dir = os.path.dirname(excel_path)

    start_dt = pd.to_datetime(config['analysis_period']['start_date'])
    end_dt = pd.to_datetime(config['analysis_period']['end_date'])
    time_label = f"{start_dt.strftime('%Y%m')}-{end_dt.strftime('%Y%m')}"

    full_output_path = os.path.join(output_dir, f"分析看板_{time_label}.html")

    if not os.path.exists(excel_path):
        print(f"错误：找不到文件 {excel_path}")
        return

    print(f"--- 正在生成可视化看板... ---")

    # 2. 获取数据
    d_interval, d_full_time, d_full_zone, d_raw_fixed = get_base_data(excel_path)

    # 3. 创建页面
    page = Page(layout=Page.SimplePageLayout)
    page.add(create_pie_component(d_interval, f"【{time_label}】月订单总金额组成", "明细项", "交易金额(元)"))
    page.add(create_line_component(d_full_time, "每月订单数量趋势", "时间/专区", "订单数量"))
    page.add(create_bar_component(d_full_time, "每月订单总金额", "时间/专区", "交易金额(元)"))
    page.add(create_pie_component(d_full_zone, "订单总金额组成", "专区/时间", "交易金额(元)", is_total_sheet=True))
    page.add(create_line_component(d_raw_fixed, "每月交易商品数量趋势", None, None, is_item_nunique=True))

    # 4. 渲染
    page.render(full_output_path)
    print(f"看板已成功生成并输出至: {full_output_path}")


if __name__ == "__main__":
    run_dashboard_output()