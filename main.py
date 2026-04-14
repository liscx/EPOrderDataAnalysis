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
    parser.add_argument("--input", help="输入原始 Excel 文件路径")
    parser.add_argument("--output", help="输出结果 Excel 文件路径")
    parser.add_argument("--start", help="分析起始日期 (YYYY-MM-DD)")
    parser.add_argument("--end", help="分析结束日期 (YYYY-MM-DD)")
    
    args = parser.parse_args()
    
    # 加载配置
    config = load_config(args.config)
    if not config:
        print(f"错误：无法加载配置文件 {args.config}")
        return
        
    config["execution_mode"] = args.mode

    # 命令行参数覆盖配置文件逻辑
    if "file_config" not in config:
        config["file_config"] = {}
        
    if args.input:
        config["file_config"]["input_file"] = args.input
        print(f"-> 命令行输入路径强制覆盖: {args.input}")

    if args.output:
        config["file_config"]["output_file"] = args.output
        print(f"-> 命令行输出路径强制覆盖: {args.output}")

    if "analysis_period" not in config:
        config["analysis_period"] = {}

    if args.start:
        config["analysis_period"]["start_date"] = args.start
        print(f"-> 命令行起始日期强制覆盖: {args.start}")

    if args.end:
        config["analysis_period"]["end_date"] = args.end
        print(f"-> 命令行结束日期强制覆盖: {args.end}")
    
    # 执行流程
    try:
        start_analysis_flow(config)
    except KeyboardInterrupt:
        print("\n用户取消执行。")
    except Exception as e:
        print(f"\n执行失败: {e}")

if __name__ == "__main__":
    main()
