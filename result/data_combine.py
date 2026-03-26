import os
import time
import base64
import pandas as pd
from docx import Document
from docx.shared import Inches
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# --- 配置区 ---
EXCEL_FILE = '数据清洗汇总结果.xlsx'
HTML_FILE = '分析看板_202602-202602.html'
TEMPLATE_FILE = '运营报告模板.docx'
OUTPUT_FILE = '3月运营报告_全量集成版.docx'

# 图表映射：Word 标记 -> HTML div id
CHART_MAPPING = {
    "{{**每月订单数量趋势**}}": "1b6e58a40b394e60857af14d5b7b8ec8",
    "{{**每月订单总金额**}}": "dc0de83a05b74081bb1f96cfa0bb85b3",
    "{{**订单总金额组成**}}": "eb93e670de6b4a5ba60bff478c1ec7bb",
    "{{**每月交易商品数量趋势**}}": "1e91645237624656a64ab2795d70e6b4",
    "{{**【202602-202602】月订单总金额组成**}}": "f63ca06ccec641a5b862f5e3dd6f5de7"
}

# --- 1. 基础工具函数 ---
def format_m(val):
    """金额格式化"""
    try:
        if pd.isna(val) or val == "" or val == "--": return "0.00"
        return f"{float(str(val).replace(',', '')):,.2f}"
    except:
        return "0.00"


def safe_split(val, expected_len):
    """严格按 | 切分"""
    if pd.isna(val) or val == "": return ["0"] * expected_len
    parts = str(val).split('|')
    if len(parts) < expected_len:
        parts.extend(["0"] * (expected_len - len(parts)))
    return parts


def fill_table_data(table, df, mapping, max_rows=None, skip_col_idx=None):
    """通用表格填充"""
    if df is None or df.empty: return
    data = df.head(max_rows) if max_rows else df
    for i, (_, row) in enumerate(data.iterrows()):
        if i + 1 >= len(table.rows): break
        cells = table.rows[i + 1].cells
        for col_idx, excel_col in mapping.items():
            if col_idx < len(cells):
                if skip_col_idx is not None and col_idx == skip_col_idx:
                    cells[col_idx].text = ""
                else:
                    val = row.get(excel_col, "")
                    if any(x in str(excel_col) for x in ["金额", "总额", "单价", "计算"]):
                        cells[col_idx].text = format_m(val)
                    else:
                        cells[col_idx].text = str(val) if pd.notna(val) else ""


# --- 2. 浏览器图表提取函数 ---
def extract_charts_as_images(html_path):
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    try:
        driver = webdriver.Chrome(options=chrome_options)
    except:
        print("❌ 浏览器驱动启动失败，请检查 chromedriver 是否在路径中。")
        return {}

    abs_url = 'file:///' + os.path.abspath(html_path).replace('\\', '/')
    driver.get(abs_url)
    time.sleep(3)  # 等待渲染

    image_paths = {}
    for mark, div_id in CHART_MAPPING.items():
        js_code = f"""
            var ins = echarts.getInstanceByDom(document.getElementById('{div_id}'));
            return ins ? ins.getDataURL({{type:'png', pixelRatio:2, backgroundColor:'#fff'}}) : null;
        """
        base64_data = driver.execute_script(js_code)
        if base64_data:
            img_bytes = base64.b64decode(base64_data.split(',')[1])
            temp_name = f"temp_{div_id}.png"
            with open(temp_name, "wb") as f: f.write(img_bytes)
            image_paths[mark] = temp_name
            print(f"📸 成功导出图表: {mark}")
    driver.quit()
    return image_paths


# --- 3. 主集成执行逻辑 ---
def main():
    print(f"🚀 开始全量集成化处理...")

    # A. 提取图表图片
    chart_images = extract_charts_as_images(HTML_FILE)

    # B. 读取 Excel 数据
    sheets = pd.read_excel(EXCEL_FILE, sheet_name=None)
    df_core = sheets.get('核心话术数据').set_index('维度')
    d = df_core.to_dict('index')

    # C. 构建话术字典
    c = {"截止日期": "2026年03月31日", "产品名称": "阳光优采"}
    v1 = safe_split(d['话术1']['数值'], 4)
    i1 = safe_split(d['话术1']['关联信息'], 4)
    c.update({
        "采购人数量": v1[0], "供应商数量": v1[1], "电商数量": v1[2], "本地供应商数量": v1[3],
        "订单数量": i1[0], "订单总额": format_m(i1[1]), "已完成订单数量": i1[2], "已完成订单总额": format_m(i1[3])
    })

    i2 = safe_split(d['话术2']['关联信息'], 9)
    c.update({
        "产生订单专区数量": d['话术2']['数值'],
        "主要专区1": i2[1], "主要专区1订单": i2[2], "主要专区1占比": i2[3],
        "主要专区2": i2[4], "主要专区2订单": i2[5], "主要专区2占比": i2[6]
    })

    v3 = safe_split(d['话术3']['数值'], 3)
    i3 = safe_split(d['话术3']['关联信息'], 10)
    cur_m, pre_m = float(i3[0]), float(i3[1])
    c.update({
        "月份": v3[0], "当月产生订单专区数量": v3[1], "当月订单数量": v3[2],
        "当月交易总额": format_m(cur_m), "上月交易总额": format_m(pre_m),
        "增长金额": format_m(cur_m - pre_m),
        "环比增长率": f"{((cur_m - pre_m) / pre_m * 100):.2f}%" if pre_m > 0 else "0.00%",
        "主要贡献专区": i3[2], "主要贡献专区订单": i3[3], "主要贡献专区总额": format_m(i3[4]),
        "次要贡献专区": i3[5], "次要贡献专区订单": i3[6], "次要贡献专区总额": format_m(i3[7])
    })

    i4 = safe_split(d['话术4']['关联信息'], 8)
    c.update({
        "活跃采购人数量": d['话术4']['数值'],
        "最高订单数": i4[0], "最高订单采购人": i4[1], "最高订单总额": format_m(i4[2]), "最高订单专区": i4[3],
        "最高金额": format_m(i4[4]), "最高金额采购人": i4[5], "最高金额订单数": i4[6], "最高金额专区": i4[7]
    })

    c.update({
        "新注册采购人数量": d['话术5']['数值'], "新注册采购人专区": d['话术5']['关联信息'],
        "供应商总数": v1[1], "商品总数": d['话术9']['数值'], "商品品类": d['话术9']['关联信息']
    })

    # D. 操作 Word
    doc = Document(TEMPLATE_FILE)

    # 1. 替换文字占位符
    for target in list(doc.paragraphs) + [cell for t in doc.tables for r in t.rows for cell in r.cells]:
        for k, v in c.items():
            if f"{{{{{k}}}}}" in target.text:
                target.text = target.text.replace(f"{{{{{k}}}}}", str(v))

    # 2. 插入高清图片
    for para in doc.paragraphs:
        for mark, img_path in chart_images.items():
            if mark in para.text:
                para.text = para.text.replace(mark, "")
                run = para.add_run()
                run.add_picture(img_path, width=Inches(6.0))

    # 3. 填充业务表格 (第3个表开始)
    ts = doc.tables
    fill_table_data(ts[2], sheets.get('采购企业汇总表'),
                    {0: '序号', 1: '采购企业', 2: '专区名称', 3: '订单数量', 4: '订单总额（元）', 5: '首次订单日期',
                     6: '末次订单日期'})
    fill_table_data(ts[3], sheets.get('采购企业订单量TOP10'),
                    {0: '序号', 1: '采购企业', 2: '专区名称', 3: '省市', 4: '订单数量', 5: '商品数量',
                     6: '订单金额（元）'}, 10)
    fill_table_data(ts[4], sheets.get('采购企业交易额TOP10'),
                    {0: '序号', 1: '采购企业', 2: '专区名称', 3: '省市', 4: '订单数量', 5: '商品数量',
                     6: '订单金额（元）'}, 10)
    fill_table_data(ts[5], sheets.get('供应商汇总表'),
                    {0: '序号', 1: '供应商', 3: '订单数量', 4: '订单总额（元）', 5: '商品数量', 6: '专区名称'},
                    skip_col_idx=2)
    fill_table_data(ts[7], sheets.get('供应商订单量TOP10'),
                    {0: '序号', 1: '单位名称', 2: '供应商类型', 3: '订单数量', 4: '商品数量', 5: '订单金额（元）'}, 10)
    fill_table_data(ts[8], sheets.get('供应商交易额TOP10'),
                    {0: '序号', 1: '单位名称', 2: '供应商类型', 3: '订单金额（元）', 4: '订单数量', 5: '商品数量'}, 10)
    fill_table_data(ts[9], sheets.get('商品销售数量TOP10'),
                    {0: '序号', 1: '商品名称', 2: '供应商', 3: '销售数量', 4: '销售总额_计算', 5: '平均单价',
                     6: '专区名称'}, 10)
    fill_table_data(ts[10], sheets.get('商品销售金额TOP10'),
                    {0: '序号', 1: '商品名称', 2: '供应商', 3: '销售数量', 4: '销售总额_计算', 5: '平均单价',
                     6: '专区名称'}, 10)

    # E. 保存并清理图片
    doc.save(OUTPUT_FILE)
    for p in chart_images.values(): os.remove(p)
    print(f"✨ 所有任务圆满完成！报告：{OUTPUT_FILE}")


if __name__ == "__main__":
    main()