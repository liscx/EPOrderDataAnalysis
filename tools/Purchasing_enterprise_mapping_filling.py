import pandas as pd
import os


def fill_in_place(target_file, mapping_file):
    # 1. 加载映射表
    try:
        mapping_df = pd.read_csv(mapping_file, encoding='utf_8_sig')
    except:
        mapping_df = pd.read_csv(mapping_file, encoding='gbk')

    mapping_dict = dict(zip(mapping_df['采购部门'].astype(str).str.strip(),
                            mapping_df['采购企业'].astype(str).str.strip()))

    # 2. 读取目标原文件
    print(f"正在读取并处理原文件: {target_file}")
    df = pd.read_excel(target_file)

    if '采购部门' not in df.columns:
        print("错误：原文件中未找到'采购部门'列")
        return

    # 3. 执行填充
    # 逻辑：如果'采购企业'原本为空，则根据'采购部门'填充；如果已有值，则保持不变（或根据需求覆盖）
    # 这里采用“如果为空则填充”的策略，如果你想强制全部按映射表更新，直接用 df['采购企业'] = ...

    def do_fill(row):
        dept = str(row['采购部门']).strip()
        # 如果当前企业名称为空(NaN)或者属于无效字符串，则从字典查找
        if pd.isna(row.get('采购企业')) or str(row.get('采购企业')).strip() == '' or str(row.get('采购企业')) == 'nan':
            return mapping_dict.get(dept, row.get('采购企业'))
        return row.get('采购企业')

    df['采购企业'] = df.apply(do_fill, axis=1)

    # 4. 写回原文件
    # 注意：Excel 此时不能被其他程序（如 Office）打开，否则会报错 Permission denied
    try:
        df.to_excel(target_file, index=False)
        print("--- 填充成功！原文件已更新 ---")
    except PermissionError:
        print("保存失败：请先关闭 Excel 程序后再运行脚本！")


# --- 设置参数 ---
target_file = '阳光优采交易订单.xlsx'  # 你的原文件路径
mapping_file = '采购部门企业对照表.csv'  # 映射字典路径

fill_in_place(target_file, mapping_file)