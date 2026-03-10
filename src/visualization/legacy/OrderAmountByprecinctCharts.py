import pandas as pd
import yaml
import os
from pyecharts import options as opts
from pyecharts.charts import Pie
from pyecharts.globals import CurrentConfig

# 资源加载设置
CurrentConfig.ONLINE_HOST = "https://assets.pyecharts.org/assets/"


def load_config():
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def generate_total_zone_pie_chart():
    # 1. 加载配置
    config = load_config()
    file_path = config['file_config']['output_file']

    if not os.path.exists(file_path):
        print(f"错误：找不到文件 {file_path}")
        return

    # 2. 读取数据
    df = pd.read_excel(file_path, sheet_name="汇总_专区_全量")
    df.columns = [c.strip() for c in df.columns]

    # 3. 数据处理：筛选“专区/时间”列中包含“小计”的行
    # 这类行代表了每个专区的历史累计总额
    zone_df = df[df['专区/时间'].str.contains('小计', na=False)].copy()

    # 清洗专区名称：去掉“ 小计”后缀，方便展示
    zone_df['专区名称'] = zone_df['专区/时间'].str.replace(' 小计', '')

    # 准备数据对
    data_pair = [
        (str(row['专区名称']), round(float(row['交易金额(元)']), 2))
        for _, row in zone_df.iterrows()
        if row['交易金额(元)'] > 0  # 仅展示有金额的专区
    ]

    if not data_pair:
        print("警告：没有发现有效的专区小计金额数据。")
        return

    # 4. 创建饼图（环形图样式）
    pie = (
        Pie(init_opts=opts.InitOpts(width="900px", height="600px", bg_color="white"))
        .add(
            series_name="历史总销售额",
            data_pair=data_pair,
            radius=["35%", "65%"],  # 环形设计，更现代
            # 设置标签：显示 专区名: 金额 (百分比)
            label_opts=opts.LabelOpts(
                formatter="{b}: {c}元\n({d}%)",
                color="#333"
            )
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(
                title="全量历史销售总额 - 专区分布",
                pos_left="center",
                title_textstyle_opts=opts.TextStyleOpts(color="#333")
            ),
            legend_opts=opts.LegendOpts(is_show=False),  # 依照要求去掉图例
            toolbox_opts=opts.ToolboxOpts(
                is_show=True,
                feature={"saveAsImage": {"title": "导出图片"}}
            )
        )
        # 设置配色方案为蓝色调系（通过自定义颜色列表实现）
        .set_colors(["#409EFF", "#67C23A", "#E6A23C", "#F56C6C", "#909399"])
    )

    # 5. 保存结果
    output_file = "全量专区金额占比饼图.html"
    pie.render(output_file)
    print(f"全量专区占比图已生成：{output_file}")


if __name__ == "__main__":
    generate_total_zone_pie_chart()