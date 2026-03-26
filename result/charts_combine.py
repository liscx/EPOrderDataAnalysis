import os
import time
import base64
import pandas as pd
from docx import Document
from docx.shared import Inches
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# --- 配置区 ---
HTML_FILE = '分析看板_202602-202602.html'
TEMPLATE_FILE = '运营报告模板.docx'
OUTPUT_FILE = '最终运营报告_含高清图.docx'

# 映射 Word 标记 到 HTML 中的 div id
CHART_MAPPING = {
    "{{**每月订单数量趋势**}}": "1b6e58a40b394e60857af14d5b7b8ec8",
    "{{**每月订单总金额**}}": "dc0de83a05b74081bb1f96cfa0bb85b3",
    "{{**订单总金额组成**}}": "eb93e670de6b4a5ba60bff478c1ec7bb",
    "{{**每月交易商品数量趋势**}}": "1e91645237624656a64ab2795d70e6b4",
    "{{**【202602-202602】月订单总金额组成**}}": "f63ca06ccec641a5b862f5e3dd6f5de7"
}


def get_driver():
    """初始化浏览器驱动，避开自动下载"""
    chrome_options = Options()
    chrome_options.add_argument('--headless')  # 无头模式
    chrome_options.add_argument('--disable-gpu')

    # 如果你的 chromedriver.exe 不在系统变量里，请取消下行注释并填入路径
    # driver = webdriver.Chrome(executable_path='C:/path/to/chromedriver.exe', options=chrome_options)

    try:
        driver = webdriver.Chrome(options=chrome_options)
        return driver
    except Exception as e:
        print(f"驱动启动失败: {e}")
        print("请确保已安装 Chrome 并在终端执行 'pip install chromedriver-binary' 或手动下载匹配的驱动")
        return None


def extract_echarts_as_images(html_path):
    """利用 pyecharts 内部 JS 接口导出图片"""
    driver = get_driver()
    if not driver: return {}

    # 获取 HTML 绝对路径
    abs_url = 'file:///' + os.path.abspath(html_path).replace('\\', '/')
    driver.get(abs_url)

    # 等待 Echarts 渲染动画（很重要，否则图会缺一块）
    time.sleep(3)

    extracted_images = {}
    for mark, div_id in CHART_MAPPING.items():
        try:
            # 执行 JS：获取该 div 对应的 echarts 实例并转为 DataURL
            # pyecharts 默认实例名通常是 chart_ + id
            js_code = f"""
                var chartIns = echarts.getInstanceByDom(document.getElementById('{div_id}'));
                if (chartIns) {{
                    return chartIns.getDataURL({{
                        type: 'png',
                        pixelRatio: 2,  // 2倍清晰度
                        backgroundColor: '#fff'
                    }});
                }} else {{
                    return null;
                }}
            """
            base64_data = driver.execute_script(js_code)

            if base64_data:
                # 提取 base64 部分并解码
                img_str = base64_data.split(',')[1]
                img_bytes = base64.b64decode(img_str)

                temp_filename = f"chart_{div_id}.png"
                with open(temp_filename, "wb") as f:
                    f.write(img_bytes)

                extracted_images[mark] = temp_filename
                print(f"✅ 成功导出图表: {mark}")
            else:
                print(f"⚠️ 未找到图表实例: {div_id}")
        except Exception as e:
            print(f"❌ 导出失败 {mark}: {e}")

    driver.quit()
    return extracted_images


def process_report():
    # 1. 抓取图片
    image_map = extract_echarts_as_images(HTML_FILE)

    if not image_map:
        print("未抓取到任何图片，请检查 HTML 文件路径或浏览器驱动。")
        return

    # 2. 插入 Word
    doc = Document(TEMPLATE_FILE)

    # 遍历段落替换标记为图片
    for para in doc.paragraphs:
        for mark, img_path in image_map.items():
            if mark in para.text:
                # 清除标记文字
                para.text = para.text.replace(mark, "")
                # 插入图片
                run = para.add_run()
                run.add_picture(img_path, width=Inches(6.0))
                print(f"已将图片插入到: {mark}")

    # 3. 保存并清理
    doc.save(OUTPUT_FILE)
    for path in image_map.values():
        if os.path.exists(path):
            os.remove(path)

    print(f"\n✨ 处理完成！最终文件：{OUTPUT_FILE}")


if __name__ == "__main__":
    process_report()