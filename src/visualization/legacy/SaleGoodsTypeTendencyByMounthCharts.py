import pandas as pd
import yaml
import os
from pyecharts import options as opts
from pyecharts.charts import Line
from pyecharts.globals import CurrentConfig

# 资源加载设置
CurrentConfig.ONLINE_HOST = "https://assets.pyecharts.org/assets/"


def load_config():
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def generate_item_type_trend():
    config = load_config()
    file_path = config['file_config']['output_file']

    if not os.path.exists(file_path):
        print("错误：未找到文件")
        return

    # 1. 读取原始明细数据
    df = pd.read_excel(file_path, sheet_name="清洗后数据")

    # --- 核心适配逻辑：修复 FutureWarning ---
    # 直接使用 .ffill() 方法替代 fillna(method='ffill')
    df['订单日期'] = df['订单日期'].ffill()
    df['专区名称'] = df['专区名称'].ffill()
    df['订单号'] = df['订单号'].ffill()

    # 2. 清洗与转换
    # 剔除商品名称为空的行（真正没有货物的行）
    df = df.dropna(subset=['商品名称'])

    # 强制转换日期格式，确保 resample 正常工作
    df['订单日期'] = pd.to_datetime(df['订单日期'], errors='coerce')
    df = df.dropna(subset=['订单日期'])

    # 3. 统计每月去重后的商品种类
    # 将日期对齐到月初
    df['月份'] = df['订单日期'].dt.to_period('M').dt.to_timestamp()

    # 分组统计去重后的商品名称数量
    trend_series = df.groupby('月份')['商品名称'].nunique()

    # 补全连续月份（MS=月初），缺失填充0
    trend_series = trend_series.resample('MS').asfreq().fillna(0)

    x_data = [d.strftime('%Y-%m') for d in trend_series.index]
    y_data = trend_series.astype(int).tolist()

    # 4. 准备静态引导线
    markline_data = [[{"coord": [x_data[i], y_data[i]]}, {"coord": [x_data[i], 0]}] for i in range(len(x_data))]

    # 5. 绘图 (实心蓝点、无图例、含引导线)
    line = (
        Line(init_opts=opts.InitOpts(width="1000px", height="500px", bg_color="white"))
        .add_xaxis(xaxis_data=x_data)
        .add_yaxis(
            series_name="商品种类",
            y_axis=y_data,
            is_smooth=False,  # 直线
            symbol="circle",  # 圆点
            symbol_size=6,  # 小尺寸
            itemstyle_opts=opts.ItemStyleOpts(color="#409EFF", border_color="#409EFF"),
            label_opts=opts.LabelOpts(is_show=True, position="top"),
            linestyle_opts=opts.LineStyleOpts(width=2, color="#409EFF"),
            markline_opts=opts.MarkLineOpts(
                data=markline_data,
                symbol=["none", "none"],
                label_opts=opts.LabelOpts(is_show=False),
                linestyle_opts=opts.LineStyleOpts(type_="dashed", color="#DCDFE6", width=1)
            )
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(title="每月交易商品种类趋势", pos_left="center"),
            legend_opts=opts.LegendOpts(is_show=False),
            tooltip_opts=opts.TooltipOpts(is_show=False),
            xaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=False), boundary_gap=True),
            yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=False)),
            toolbox_opts=opts.ToolboxOpts(
                is_show=True,
                feature={"saveAsImage": {"title": "导出图片"}}
            )
        )
    )

    line.render("每月商品种类趋势图_最终版.html")
    print("代码已修正，警告已消除，趋势图生成成功。")


if __name__ == "__main__":
    generate_item_type_trend()