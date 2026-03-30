import os
import sys
import yaml
import pandas as pd
from docx import Document

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def load_config():
    base_dir = get_base_dir()
    config_path = os.path.join(base_dir, "config.yaml")
    if not os.path.exists(config_path):
        config_path = "config.yaml"
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def get_paths():
    config = load_config()
    start_dt = pd.to_datetime(config['analysis_period']['start_date'])
    end_dt = pd.to_datetime(config['analysis_period']['end_date'])
    time_label = str(end_dt.month)
    
    excel_path = config.get('file_config', {}).get('output_file', '')
    output_dir = os.path.dirname(excel_path) if excel_path else os.path.join(get_base_dir(), 'result')
    
    EXCEL_FILE = excel_path
    TEMP_CHART_FILE = os.path.join(output_dir, f"{end_dt.month}月运营报告_中间版.docx")
    if os.path.exists(TEMP_CHART_FILE):
        TEMPLATE_FILE = TEMP_CHART_FILE
    else:
        TEMPLATE_FILE = os.path.join(output_dir, f"{time_label}月运营报告_gen.docx")
        
    OUTPUT_FILE = os.path.join(output_dir, f"{time_label}月运营报告_gen.docx")
    return EXCEL_FILE, TEMPLATE_FILE, OUTPUT_FILE

# --- 1. 基础工具函数 ---
def format_m(val):
    """金额格式化"""
    try:
        if pd.isna(val) or val == "" or val == "--": return "0.00"
        return f"{float(str(val).replace(',', '')):,.2f}"
    except:
        return "0.00"


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


# --- 2. 主集成执行逻辑 ---
def main():
    print(f"🚀 开始全量集成化处理...")
    
    # 动态载入当前系统的时文路径和配置设定
    EXCEL_FILE, TEMPLATE_FILE, OUTPUT_FILE = get_paths()

    # A. 读取 Excel 数据
    sheets = pd.read_excel(EXCEL_FILE, sheet_name=None)
    df_core = sheets.get('核心提取数据')
    c_loaded = dict(zip(df_core['占位符'], df_core['填充值']))

    # B. 构建话术字典
    c = {
        "{{截止日期}}": "2026年03月31日",
        "{{产品名称}}": "阳光优采"
    }
    c.update(c_loaded)

    # C. 操作 Word
    doc = Document(TEMPLATE_FILE)

    # 1. 替换文字占位符 (优化性能版本)
    for target in list(doc.paragraphs) + [cell for t in doc.tables for r in t.rows for cell in r.cells]:
        orig_text = target.text
        if not orig_text or '{{' not in orig_text:
            continue
            
        new_text = orig_text
        for k, v in c.items():
            if str(k) in new_text:
                new_text = new_text.replace(str(k), str(v))
                
        if new_text != orig_text:
            target.text = new_text

    # 2. 填充业务表格 (第3个表开始)
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
                    {0: '序号', 1: '供应商', 2: '供应商类型', 3: '订单数量', 4: '订单总额（元）', 5: '商品数量', 6: '专区名称'})
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

    # D. 保存生成的报告
    doc.save(OUTPUT_FILE)
    print(f"\n✨ 所有任务圆满完成！报告：{OUTPUT_FILE}")


if __name__ == "__main__":
    main()