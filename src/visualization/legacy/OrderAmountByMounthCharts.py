import pandas as pd
import yaml
import os
from pyecharts import options as opts
from pyecharts.charts import Bar
from pyecharts.globals import CurrentConfig

# 静态资源设置，防止 H5 空白
CurrentConfig.ONLINE_HOST = "https://assets.pyecharts.org/assets/"


def load_config():
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def generate_amount_bar_chart():
    # 1. 加载配置
    config = load_config()
    file_path = config['file_config']['output_file']

    if not os.path.exists(file_path):
        print(f"错误：找不到文件 {file_path}")
        return

    # 2. 读取数据
    df = pd.read_excel(file_path, sheet_name="汇总_时间_全量")
    df.columns = [c.strip() for c in df.columns]

    # 3. 数据处理：提取小计行并补全月份
    # 筛选包含“小计”的行
    bar_df = df[df['时间/专区'].str.contains('小计', na=False)].copy()

    # 转换日期格式
    bar_df['日期'] = pd.to_datetime(
        bar_df['时间/专区'].str.replace(' 小计', ''),
        format='%Y年%m月'
    )

    # 核心：重采样补全连续月份，缺失金额填充 0
    bar_df = bar_df.set_index('日期').resample('MS').asfreq().fillna(0)

    x_data = [d.strftime('%Y-%m') for d in bar_df.index]
    y_data = [round(float(v), 2) for v in bar_df['交易金额(元)']]

    # 4. 生成柱状图
    bar = (
        Bar(init_opts=opts.InitOpts(width="1000px", height="500px", bg_color="white"))
        .add_xaxis(xaxis_data=x_data)
        .add_yaxis(
            series_name="交易金额",
            y_axis=y_data,
            category_gap="40%",  # 柱子之间的间距
            label_opts=opts.LabelOpts(is_show=True, position="top", formatter="{c}"),
            itemstyle_opts=opts.ItemStyleOpts(
                color="#409EFF",  # 统一使用蓝色实心
                opacity=0.8  # 稍微增加透明度，看起来更高级
            )
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(title="每月订单总金额趋势", pos_left="center"),
            legend_opts=opts.LegendOpts(is_show=False),  # 去掉图例
            tooltip_opts=opts.TooltipOpts(is_show=True, trigger="axis", axis_pointer_type="shadow"),
            # 去掉 Grid 网格线
            xaxis_opts=opts.AxisOpts(
                name="月份",
                splitline_opts=opts.SplitLineOpts(is_show=False),
            ),
            yaxis_opts=opts.AxisOpts(
                name="金额 (元)",
                splitline_opts=opts.SplitLineOpts(is_show=False),
            ),
            toolbox_opts=opts.ToolboxOpts(
                is_show=True,
                feature={"saveAsImage": {"title": "导出图片"}}
            )
        )
    )

    # 5. 保存结果
    output_file = "每月金额分布柱状图.html"
    bar.render(output_file)
    print(f"柱状图已生成：{output_file}")


if __name__ == "__main__":
    generate_amount_bar_chart()