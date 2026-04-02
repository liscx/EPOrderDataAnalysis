import time
import traceback
import sys
import os

# 挂载 src 路径以引入移动后的集成脚本
project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
src_inner = os.path.join(project_root, "src")
if src_inner not in sys.path:
    sys.path.insert(0, src_inner)

from ..result.data_combine import main as run_data_combine
from ..result.charts_combine import build_report as run_charts_combine

from ..visualization import run_report_output, run_dashboard_output
from . import (
    do_wash, 
    run_enhanced_metrics, 
    run_total_order, 
    run_purchaser_analysis,
    run_range_purchaser_top10, 
    run_supplier_analysis, 
    run_range_supplier_top10, 
    run_product_analysis
)

def start_analysis_flow(config):
    """
    由 GUI 调用的自动化分析工作流
    """
    mode = config.get("execution_mode", "all")

    print("=" * 50)
    print(f"  商品交易数据自动化分析系统启动")
    print(f"  分析区间: {config['analysis_period']['start_date']} 至 {config['analysis_period']['end_date']}")
    print("=" * 50)

    if mode == "all":
        start_time = time.time()
        try:
            print("\n" + "*"*30)
            print("[预处理] 数据完整性审计启动...")
            input_path = config['file_config']['input_path'] if 'input_path' in config['file_config'] else config['file_config']['input_file']
            output_path = config['file_config']['output_file']
            
            import pandas as pd
            df_init = pd.read_excel(input_path)
            # 预清理列名空格
            df_init.columns = [str(c).strip() for c in df_init.columns]
            
            print(f"-> 原始列名清单: {df_init.columns.tolist()}")
            
            # 检测核心列状态
            needs_fill = False
            if '采购企业' not in df_init.columns:
                print("-> ⚠️ 预警：原始数据中缺少 [采购企业] 列！")
                needs_fill = True
            else:
                nan_count = df_init['采购企业'].isna().sum()
                print(f"-> [采购企业] 空值统计: {nan_count} 行 (总计 {len(df_init)} 行)")
                if nan_count > 0:
                    needs_fill = True
            
            if needs_fill:
                from ..tools.Purchasing_enterprise_mapping_filling import fill_in_place
                from ..core.constants import MAPPING_SCAN_DIRS, MAPPING_CSV_NAME
                
                # 智能寻找映射文件
                mapping_file = None
                from ..core.config import get_base_dir
                base_dir = get_base_dir()
                for d in MAPPING_SCAN_DIRS:
                    target = os.path.join(base_dir, d, MAPPING_CSV_NAME)
                    if os.path.exists(target):
                        mapping_file = target
                        break
                
                if mapping_file:
                    print(f"-> 正在调用映射表进行回填: [{MAPPING_CSV_NAME}]")
                    fill_in_place(input_path, mapping_file)
                    print("-> 原始数据回填补全已完成。")
                else:
                    print(f"-> ⚠️ 警告：未找到映射表 [{MAPPING_CSV_NAME}]，跳过自动补全。")
            else:
                print("-> 原始数据完整性良好。")
            print("*"*30 + "\n")

            print("[1/10] 正在执行数据清洗与区间切片...")
            do_wash()

            print("\n[2/10] 正在执行月度/专区交易明细统计...")
            run_total_order()

            print("\n[3/10] 正在执行采购企业历史全量汇总...")
            run_purchaser_analysis()

            print("\n[4/10] 正在执行指定区间采购企业 Top10 排行...")
            run_range_purchaser_top10()

            print("\n[5/10] 正在执行供应商历史全量汇总...")
            run_supplier_analysis()

            print("\n[6/10] 正在执行指定区间供应商 Top10 排行...")
            run_range_supplier_top10()

            print("\n[7/12] 正在执行商品维度分析...")
            run_product_analysis()

            print("\n[8/12] 正在提取特定供应商原始数据...")
            from ..extra.new_supplier_data_Prasing import run_new_supplier_parsing
            run_new_supplier_parsing()

            print("\n[9/12] 正在提取全量核心概览指标...")
            run_enhanced_metrics()
            
            print("\n[10/12] 正在创作可视化看板...")
            run_dashboard_output()
            
            print("\n[11/12] 正在全量集成终极运营报告数据...")
            run_data_combine()
            
            print("\n[12/12] 正在植入高清图表至报告...")
            run_charts_combine()




            end_time = time.time()
            print("\n" + "=" * 50)
            print(f"  所有分析任务已成功完成！")
            print(f"  总耗时: {end_time - start_time:.2f} 秒")
            print("=" * 50)

        except Exception as e:
            print(f"\n[运行异常]: {e}")
            traceback.print_exc()
    else:
        print(f"提示：当前模式为 '{mode}'，跳过自动化流程。")
