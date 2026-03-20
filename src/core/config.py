import yaml
import os

def load_config(config_path="config.yaml"):
    """
    统一的配置文件加载器
    """
    if not os.path.exists(config_path):
        # 尝试从根目录寻找
        base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        config_path = os.path.join(base_dir, "config.yaml")
        
    if not os.path.exists(config_path):
         return {}

    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)
