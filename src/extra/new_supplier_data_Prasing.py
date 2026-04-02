import re
import pandas as pd
import os
import sys

from ..core import get_base_dir

def run_new_supplier_parsing():
    """工作流集成入口：提取供应商原始数据并写入结果表"""
    from ..core import load_config
    config = load_config()
    
    # 1. 定位输入的 TXT 文件（支持打包环境下平级读取）
    txt_filename = "GYSdata.txt"
    base_dir = get_base_dir()
    
    # 从核心常量库载入扫描路径规则
    from ..core.constants import GYS_SCAN_DIRS
    paths_to_try = [os.path.join(base_dir, d, txt_filename).replace("\\\\", "\\") for d in GYS_SCAN_DIRS]
    # 对当前工作目录做额外兜底
    paths_to_try.append(os.path.join(os.getcwd(), txt_filename))
    
    file_path = None
    for p in paths_to_try:
        if os.path.exists(p):
            file_path = p
            break
            
    if not file_path:
        print(f"警告：未找到供应商原始数据文件 {txt_filename}，将跳过此步。")
        return

    # 2. 读取配置信息（含输出路径与时间区间）
    excel_path = config['file_config']['output_file']
    period = config.get('analysis_period', {})
    start_str = period.get('start_date', '2000-01-01')
    end_str = period.get('end_date', '2099-12-31')
    
    # 转换为日期对象以便后续过滤
    try:
        start_date = pd.to_datetime(start_str)
        # 确保包含结束当天的全部数据
        end_date = pd.to_datetime(end_str) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    except:
        start_date = pd.to_datetime('2000-01-01')
        end_date = pd.to_datetime('2099-12-31')

    # 3. 核心提取逻辑
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    pattern = re.compile(
        r'"danweiname":\s*"([^"]+?)".*?'
        r'"rukudate":\s*"([^"]+?)".*?'
        r'"tenantname":\s*"([^"]+?)"',
        re.DOTALL
    )

    matches = pattern.findall(content)
    if not matches:
        print(f"未在 {txt_filename} 中解析到供应商数据。")
        return

    data_list = []      # 存储区间过滤后的数据 (start_date ~ end_date)
    all_data_list = []  # 存储截至 end_date 的全量历史数据 (<= end_date)
    
    for name, date_str, tenant in matches:
        try:
            # 过滤日期逻辑
            cur_date = pd.to_datetime(date_str.strip())
            
            # 逻辑 1：提取截至 end_date 的所有数据 (用户新需求)
            if cur_date <= end_date:
                all_data_list.append({
                    "单位名称": name.strip(),
                    "专区名称": tenant.strip(),
                    "入库日期": date_str.strip()
                })
            
            # 逻辑 2：提取在指定区间 [start_date, end_date] 内的数据 (原逻辑)
            if start_date <= cur_date <= end_date:
                data_list.append({
                    "单位名称": name.strip(),
                    "专区名称": tenant.strip(),
                    "入库日期": date_str.strip()
                })
        except:
            continue

    if not all_data_list:
        print(f"在 {end_str} 之前未找到任何符合条件的供应商数据。")
        return

    # 4. 写入总表的不同 Sheet
    try:
        # 判断文件是否存在以决定写入模式
        if os.path.exists(excel_path):
            writer = pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace')
        else:
            writer = pd.ExcelWriter(excel_path, engine='openpyxl')

        with writer:
            # 写入全量数据 (截至 end_date)
            df_all = pd.DataFrame(all_data_list)
            df_all.insert(0, '序号', range(1, len(df_all) + 1))
            sheet_all = '供应商截至数据'
            df_all.to_excel(writer, sheet_name=sheet_all, index=False)
            
            # 写入区间数据 (start_date 到 end_date)
            if data_list:
                df_range = pd.DataFrame(data_list)
                df_range.insert(0, '序号', range(1, len(df_range) + 1))
                sheet_range = '供应商信息提取'
                df_range.to_excel(writer, sheet_name=sheet_range, index=False)
                print(f"成功更新 Excel: [{sheet_all}] (共{len(df_all)}条) 和 [{sheet_range}] (区间共{len(df_range)}条)。")
            else:
                print(f"成功更新 Excel: [{sheet_all}] (共{len(df_all)}条)；区间 [{start_str} 至 {end_str}] 内无数据。")

    except Exception as e:
        print(f"集成供应商数据至 Excel 失败: {e}")


if __name__ == "__main__":
    run_new_supplier_parsing()
