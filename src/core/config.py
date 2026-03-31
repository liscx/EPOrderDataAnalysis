import sys
import os
import yaml

def get_base_dir():
    """获取程序根目录，兼容源码运行与 PyInstaller 打包后的 EXE"""
    if getattr(sys, 'frozen', False):
        # 打包后的运行环境，返回 .exe 文件所在的文件夹
        return os.path.dirname(sys.executable)
    # 源码环境：abspath(__file__) 是 src/core/config.py，向前退两层到根目录
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def load_config(config_path=None):
    """
    统一的、兼容打包的配置文件加载器
    """
    base_dir = get_base_dir()
    if config_path is None:
        config_path = os.path.join(base_dir, "config.yaml")
        
    if not os.path.exists(config_path):
         # 如果没找到，尝试在当前 CWD 辅助查找
         if os.path.exists("config.yaml"):
             config_path = "config.yaml"
         else:
             return {} # 返回空配置防止崩溃

    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)
