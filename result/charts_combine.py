import os
import time
import base64
import yaml
import pandas as pd
from docx import Document
from docx.shared import Inches
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

def get_base_dir():
    import sys
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    curr_dir = os.path.dirname(os.path.abspath(__file__))
    if os.path.basename(curr_dir) == 'result':
        return os.path.dirname(curr_dir)
    return curr_dir

def load_config():
    base_dir = get_base_dir()
    config_path = os.path.join(base_dir, "config.yaml")
    if not os.path.exists(config_path):
        config_path = "config.yaml"
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

# --- 配置区自动计算 ---
config = load_config()
end_dt = pd.to_datetime(config['analysis_period']['end_date'])
time_label = str(end_dt.month)

excel_path = config.get('file_config', {}).get('output_file', '')
output_dir = os.path.dirname(excel_path) if excel_path else os.path.join(get_base_dir(), 'result')

HTML_FILE = os.path.join(output_dir, f"分析看板_{time_label}.html")
TEMPLATE_FILE = os.path.join(get_base_dir(), '运营报告模板.docx')
OUTPUT_FILE = os.path.join(output_dir, f"{time_label}月运营报告_gen.docx")

# 这里填入 Word 中的占位符和 HTML 中对应的 div id
# 注意，其中一项使用了基于 config 的动态月份占位，来完美兼容你在文档中的动态月份
CHART_MAPPING = {
    "{{**每月订单数量趋势**}}": "1b6e58a40b394e60857af14d5b7b8ec8",
    "{{**每月订单总金额**}}": "dc0de83a05b74081bb1f96cfa0bb85b3",
    "{{**订单总金额组成**}}": "eb93e670de6b4a5ba60bff478c1ec7bb",
    "{{**每月交易商品数量趋势**}}": "1e91645237624656a64ab2795d70e6b4",
    f"{{**月订单总金额组成**}}": "f63ca06ccec641a5b862f5e3dd6f5de7"
}

def get_charts_via_js(html_path):
    """通过执行JS脚本直接从Canvas获取Base64图片，不依赖物理截图"""
    print("-> 正在初始化 Chrome 浏览器...")
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--allow-file-access-from-files')
    chrome_options.add_argument('--disable-web-security')
    
    # 注意：为了避免 webdriver_manager 的网络报错，请确保你本地有 chromedriver 
    # 或者直接指定你电脑上浏览器的路径
    driver = webdriver.Chrome(options=chrome_options)
    
    abs_path = 'file:///' + os.path.abspath(html_path).replace('\\', '/')
    
    try:
        driver.set_page_load_timeout(30)
        driver.get(abs_path)
        print("-> 正在等待 Echarts 就绪...")
        # 轮询探测 Echarts 库及图表实例是否渲染好
        for i in range(15):
            has_echarts = driver.execute_script("return typeof echarts !== 'undefined';")
            if has_echarts:
                time.sleep(3)  # 找到之后再等待 3 秒完整执行图表绘制动画
                break
            time.sleep(1)
    except Exception as e:
        print(f"❌ 页面加载失败: {e}")
        driver.quit()
        return {}

    images = {}
    for mark, div_id in CHART_MAPPING.items():
        try:
            # 这里的 JS 逻辑相当于手动点击了导出按钮并拦截了返回的数据
            js_script = f"""
                if (typeof echarts === 'undefined') return null;
                var dom = document.getElementById('{div_id}');
                if (dom) {{
                    var chart = echarts.getInstanceByDom(dom);
                    if (chart) {{
                        return chart.getDataURL({{
                            type: 'png',
                            pixelRatio: 2,
                            backgroundColor: '#fff'
                        }});
                    }}
                }}
                return null;
            """
            base64_str = driver.execute_script(js_script)
            
            if base64_str:
                # 去掉 base64 前缀并保存
                img_data = base64.b64decode(base64_str.split(',')[1])
                temp_path = f"temp_{div_id}.png"
                with open(temp_path, "wb") as f:
                    f.write(img_data)
                images[mark] = temp_path
                print(f"✅ 已提取高清图表: {mark}")
            else:
                 print(f"⚠️ 无法提取图表: {mark} (未找到实例或未渲染完毕)")
        except Exception as e:
            print(f"❌ 提取失败 {mark}: {e}")

    driver.quit()
    return images

def build_report():
    print(f"开始生成报告，读取数据看板: {HTML_FILE}")
    if not os.path.exists(HTML_FILE):
        print(f"❌ 找不到 HTML 看板文件: {HTML_FILE}")
        return

    # 1. 提取图片
    image_paths = get_charts_via_js(HTML_FILE)
        
    # 2. 打开文档并替换
    doc = Document(TEMPLATE_FILE)
    
    if image_paths:
        # 替换图片标记
        for para in doc.paragraphs:
            for mark, path in image_paths.items():
                if mark in para.text:
                    para.text = para.text.replace(mark, "")
                    run = para.add_run()
                    run.add_picture(path, width=Inches(6.0))
                    try:
                        os.remove(path) # 插入后删除临时文件
                    except OSError:
                        pass
    else:
        print("没有抓取到任何图片，将仅复制原模板以便后续数据填充。")

    # (此处可插入你之前的表格填充逻辑...)
    
    doc.save(OUTPUT_FILE)
    print(f"🚀 图表处理成功！中间报告存放于: {OUTPUT_FILE}")

if __name__ == "__main__":
    build_report()