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
    from .constants import CONFIG_SCAN_DIRS
    base_dir = get_base_dir()
    
    if config_path is None:
        filename = "config.yaml"
        # 尝试遍历所有扫描路径以定位配置文件
        for d in CONFIG_SCAN_DIRS:
            target_path = os.path.join(base_dir, d, filename)
            if os.path.exists(target_path):
                config_path = target_path
                break
        
        # 兜底：如果依然找不到，尝试在当前工作目录寻找
        if not config_path or not os.path.exists(config_path):
            if os.path.exists("config.yaml"):
                config_path = "config.yaml"
            else:
                return {} # 返回空配置防止崩溃

    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}
