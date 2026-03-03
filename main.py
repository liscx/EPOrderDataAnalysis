import yaml
import os
import time

# 导入重构后的各个模块函数
from TestDataWash import do_wash

from CoreMetricsExtraction import run_core_metrics
from TotalOrderAnalysis import run_analysis as run_total_order
from TotalPurchaserAnalysis import run_purchaser_analysis
from DataRangePurchaserAnalysis import run_range_purchaser_top10
from TotalSuppliersAnalysis import run_supplier_analysis
from DataRangeSuppliersAnalysis import run_range_supplier_top10
from DataRangeGoodsAnalysis import run_product_analysis


def load_config():
    """加载配置文件"""
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def main():
    # 1. 加载配置
    try:
        config = load_config()
    except Exception as e:
        print(f"错误：读取 config.yaml 失败！{e}")
        return

    mode = config.get("execution_mode", "all")

    print("=" * 60)
    print(f"  商品交易数据自动化分析系统启动  ")
    print(f"  当前模式: {mode}")
    print("=" * 60)

    # 2. 按序执行任务列表
    # 顺序：清洗 -> 全局汇总 -> 企业全量 -> 企业区间Top10 -> 供应商全量 -> 供应商区间Top10 -> 商品分析
    if mode == "all":
        start_time = time.time()

        try:
            # Step 1: 数据清洗与区间数据生成
            print("\n[Step 1/8] 正在执行数据清洗与区间切片...")
            do_wash()

            # Step 2: 交易汇总分析（时间/专区维度）
            print("\n[Step 2/8] 正在执行月度/专区交易明细统计...")
            run_total_order()

            # Step 3: 采购企业全量分析
            print("\n[Step 3/8] 正在执行采购企业历史全量汇总...")
            run_purchaser_analysis()

            # Step 4: 采购企业区间 Top10
            print("\n[Step 4/8] 正在执行指定区间采购企业 Top10 排行...")
            run_range_purchaser_top10()

            # Step 5: 供应商全量分析
            print("\n[Step 5/8] 正在执行供应商历史全量汇总...")
            run_supplier_analysis()

            # Step 6: 供应商区间 Top10
            print("\n[Step 6/8] 正在执行指定区间供应商 Top10 排行...")
            run_range_supplier_top10()

            # Step 7: 商品维度分析
            print("\n[Step 7/8] 正在执行指定区间商品交易分析与 Top10...")
            run_product_analysis()

            print("\n[Step 8/8] 正在提取全量核心概览指标...")
            run_core_metrics()

            end_time = time.time()
            print("\n" + "=" * 60)
            print(f"  所有分析任务已成功完成！")
            print(f"  总耗时: {end_time - start_time:.2f} 秒")
            print(f"  最终结果已保存至: {config['file_config']['output_file']}")
            print("=" * 60)

        except Exception as e:
            print(f"\n[运行异常] 流程在执行过程中中断：{e}")
            import traceback
            traceback.print_exc()

    else:
        print(f"提示：当前 execution_mode 为 '{mode}'，未设置为 'all'，程序不执行任何自动化流程。")


if __name__ == "__main__":
    main()