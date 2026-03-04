import pandas as pd
import yaml
import os
from pyecharts import options as opts
from pyecharts.charts import Line
from pyecharts.globals import CurrentConfig

# 静态资源设置
CurrentConfig.ONLINE_HOST = "https://assets.pyecharts.org/assets/"


def load_config():
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def generate_trend_chart():
    config = load_config()
    file_path = config['file_config']['output_file']

    if not os.path.exists(file_path): return

    # 1. 数据处理
    df = pd.read_excel(file_path, sheet_name="汇总_时间_全量")
    df.columns = [c.strip() for c in df.columns]
    trend_df = df[df['时间/专区'].str.contains('小计', na=False)].copy()

    trend_df['日期'] = pd.to_datetime(
        trend_df['时间/专区'].str.replace(' 小计', ''),
        format='%Y年%m月'
    )
    # 保证月份连续，缺失补0
    trend_df = trend_df.set_index('日期').resample('MS').asfreq().fillna(0)

    x_data = [d.strftime('%Y-%m') for d in trend_df.index]
    y_data = trend_df['订单数量'].astype(int).tolist()

    # 2. 构建静态垂直引导线
    markline_data = []
    for i in range(len(x_data)):
        markline_data.append([
            {"coord": [x_data[i], y_data[i]]},
            {"coord": [x_data[i], 0]}
        ])

    # 3. 生成折线图
    line = (
        Line(init_opts=opts.InitOpts(width="1000px", height="500px", bg_color="white"))
        .add_xaxis(xaxis_data=x_data)
        .add_yaxis(
            series_name="订单数量",
            y_axis=y_data,
            is_smooth=False,  # 直线
            symbol="circle",  # 圆形节点
            symbol_size=6,  # --- 修改：稍小一点的尺寸 ---
            itemstyle_opts=opts.ItemStyleOpts(
                color="#409EFF",  # --- 修改：实心蓝色 ---
                border_color="#409EFF",
                border_width=1
            ),
            label_opts=opts.LabelOpts(is_show=True, position="top", color="#333"),
            linestyle_opts=opts.LineStyleOpts(width=2, color="#409EFF"),
            # 静态引导线
            markline_opts=opts.MarkLineOpts(
                data=markline_data,
                symbol=["none", "none"],
                label_opts=opts.LabelOpts(is_show=False),
                linestyle_opts=opts.LineStyleOpts(type_="dashed", color="#DCDFE6", width=1)
            )
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(title="每月订单数量趋势", pos_left="center"),
            legend_opts=opts.LegendOpts(is_show=False),  # --- 修改：去掉图例 ---
            tooltip_opts=opts.TooltipOpts(is_show=False),
            xaxis_opts=opts.AxisOpts(
                splitline_opts=opts.SplitLineOpts(is_show=False),
                boundary_gap=True
            ),
            yaxis_opts=opts.AxisOpts(
                splitline_opts=opts.SplitLineOpts(is_show=False),
            ),
            toolbox_opts=opts.ToolboxOpts(
                is_show=True,
                feature={"saveAsImage": {"title": "导出图片"}}
            )
        )
    )

    line.render("每月订单趋势图_最终简洁版.html")
    print("趋势图已生成：实心蓝点、无图例、含静态引导线。")
订单金额

if __name__ == "__main__":
    generate_trend_chart()