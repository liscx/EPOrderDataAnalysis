import time
import traceback
# 1. 导入你的业务模块
from TestDataWash import do_wash
from CoreMetricsExtraction import run_core_metrics
from TotalOrderAnalysis import run_analysis as run_total_order
from TotalPurchaserAnalysis import run_purchaser_analysis
from DataRangePurchaserAnalysis import run_range_purchaser_top10
from TotalSuppliersAnalysis import run_supplier_analysis
from DataRangeSuppliersAnalysis import run_range_supplier_top10
from DataRangeGoodsAnalysis import run_product_analysis

# 2. 导入可视化模块
from AutoCharts import build_dashboard

def start_analysis_flow(config):
    """
    由 GUI 调用的主入口函数
    """
    mode = config.get("execution_mode", "all")

    print("=" * 50)
    print(f"  商品交易数据自动化分析系统启动")
    print(f"  分析区间: {config['analysis_period']['start_date']} 至 {config['analysis_period']['end_date']}")
    print("=" * 50)

    if mode == "all":
        start_time = time.time()
        try:
            print("\n[1/9] 正在执行数据清洗与区间切片...")
            do_wash()

            print("\n[2/9] 正在执行月度/专区交易明细统计...")
            run_total_order()

            print("\n[3/9] 正在执行采购企业历史全量汇总...")
            run_purchaser_analysis()

            print("\n[4/9] 正在执行指定区间采购企业 Top10 排行...")
            run_range_purchaser_top10()

            print("\n[5/9] 正在执行供应商历史全量汇总...")
            run_supplier_analysis()

            print("\n[6/9] 正在执行指定区间供应商 Top10 排行...")
            run_range_supplier_top10()

            print("\n[7/9] 正在执行商品维度分析...")
            run_product_analysis()

            print("\n[8/9] 正在提取全量核心概览指标...")
            run_core_metrics()

            # 3. 这里是调用点：执行可视化看板生成
            build_dashboard(config)

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