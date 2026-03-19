import time
import traceback
import os

# 1. 导入业务模块
from TestDataWash import do_wash
from CoreMetricsExtraction import run_enhanced_metrics
from TotalOrderAnalysis import run_analysis as run_total_order
from TotalPurchaserAnalysis import run_purchaser_analysis
from DataRangePurchaserAnalysis import run_range_purchaser_top10
from TotalSuppliersAnalysis import run_supplier_analysis
from DataRangeSuppliersAnalysis import run_range_supplier_top10
from DataRangeGoodsAnalysis import run_product_analysis
from DescribeTemplate import run_report_output
from AutoCharts import run_dashboard_output


def start_analysis_flow(config):
    """
    由 GUI 调用的主入口函数 - 增强报错处理版
    """
    mode = config.get("execution_mode", "all")

    print("=" * 50)
    print(f"  商品交易数据自动化分析系统启动")
    print(f"  分析区间: {config['analysis_period']['start_date']} 至 {config['analysis_period']['end_date']}")
    print("=" * 50)

    if mode == "all":
        start_time = time.time()

        # 任务定义：(任务名称, 函数引用)
        tasks = [
            ("数据清洗与区间切片", do_wash),
            ("月度/专区交易明细统计", run_total_order),
            ("采购主体全量汇总 (企业/部门)", run_purchaser_analysis),
            ("指定区间采购主体 Top10", run_range_purchaser_top10),
            ("供应商全量汇总", run_supplier_analysis),
            ("指定区间供应商 Top10", run_range_supplier_top10),
            ("商品维度分析", run_product_analysis),
            ("提取全量核心概览指标", run_enhanced_metrics),
            ("创作可视化看板", run_dashboard_output),
            ("编辑月报核心话术", run_report_output)
        ]

        success_count = 0
        for i, (name, func) in enumerate(tasks, 1):
            try:
                print(f"\n[{i}/{len(tasks)}] 正在执行{name}...")
                func()
                success_count += 1
                print(f"  √ {name} 执行成功")
            except KeyError as e:
                print(f"  × [字段缺失]: 在'{name}'步骤中找不到关键列 {e}。请检查原始文件表头。")
            except FileNotFoundError as e:
                print(f"  × [文件缺失]: {e}")
            except Exception as e:
                print(f"  × [执行异常]: '{name}' 运行失败。详情: {e}")
                traceback.print_exc()

        end_time = time.time()
        print("\n" + "=" * 50)
        print(f"  所有分析任务执行完毕！")
        print(f"  成功: {success_count} | 失败: {len(tasks) - success_count}")
        print(f"  总耗时: {end_time - start_time:.2f} 秒")
        print("=" * 50)
    else:
        print(f"提示：当前模式为 '{mode}'，跳过自动化流程。")