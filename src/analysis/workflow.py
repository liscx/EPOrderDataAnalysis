import time
import traceback
import sys
import os

# 挂载 result 路径以引入新版增强集成脚本
project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if project_root not in sys.path:
    sys.path.insert(0, project_root)
from result.data_combine import main as run_data_combine
from result.charts_combine import build_report as run_charts_combine

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
            print("\n[1/10] 正在执行数据清洗与区间切片...")
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

            print("\n[7/10] 正在执行商品维度分析...")
            run_product_analysis()

            print("\n[8/10] 正在提取全量核心概览指标...")
            run_enhanced_metrics()
            
            print("\n[9/10] 正在创作可视化看板...")
            run_dashboard_output()

            print("\n[10/11] 正在植入高清图表至报告...")
            run_charts_combine()

            print("\n[11/11] 正在全量集成终极运营报告数据...")
            run_data_combine()

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
