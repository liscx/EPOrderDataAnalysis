import sys
import os
import argparse

# 确保项目根目录在 sys.path 中
root_dir = os.path.dirname(os.path.abspath(__file__))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

from src.core import load_config
from src.analysis import start_analysis_flow

def main():
    parser = argparse.ArgumentParser(description="Data Engine 命令行分析工具")
    parser.add_argument("--config", default="config.yaml", help="配置文件路径")
    parser.add_argument("--mode", choices=["all", "custom"], default="all", help="执行模式")
    
    args = parser.parse_args()
    
    # 加载配置
    config = load_config(args.config)
    if not config:
        print(f"错误：无法加载配置文件 {args.config}")
        return
        
    config["execution_mode"] = args.mode
    
    # 执行流程
    try:
        start_analysis_flow(config)
    except KeyboardInterrupt:
        print("\n用户取消执行。")
    except Exception as e:
        print(f"\n执行失败: {e}")

if __name__ == "__main__":
    main()
