import pandas as pd


def extract_department_mapping_xlsx(file_path, output_path):
    print(f"正在处理文件: {file_path}...")

    # 1. 读取 Excel 文件
    # 注意：如果运行报错提示缺少 openpyxl，请在终端执行: pip install openpyxl
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"读取失败: {e}")
        return

    # 2. 提取相关列并清理
    # 确保 '采购部门' 和 '采购企业' 存在
    required_cols = ['采购部门', '采购企业']
    if not all(col in df.columns for col in required_cols):
        print(f"错误：文件中找不到列 {required_cols}")
        print(f"当前文件包含的列名有: {df.columns.tolist()}")
        return

    # 只保留这两列，并删除任一列为空的行
    mapping_df = df[required_cols].dropna()

    # 3. 去除字符串前后空格（防止因为空格导致匹配失败）
    mapping_df['采购部门'] = mapping_df['采购部门'].astype(str).str.strip()
    mapping_df['采购企业'] = mapping_df['采购企业'].astype(str).str.strip()

    # 4. 冲突检查（同一部门对应多家企业）
    counts = mapping_df.groupby('采购部门')['采购企业'].nunique()
    conflicts = counts[counts > 1]
    if not conflicts.empty:
        print("\n--- 预警：以下部门在多家企业中出现过 ---")
        for dept in conflicts.index:
            corps = mapping_df[mapping_df['采购部门'] == dept]['采购企业'].unique()
            print(f"部门 [{dept}] 对应的企业有: {corps}")
        print("---------------------------------------\n")

    # 5. 去重：每个部门只保留一个对应的企业（默认保留第一条记录）
    mapping_df = mapping_df.drop_duplicates(subset=['采购部门'])

    # 6. 保存为新的 CSV 映射表，方便后续调用
    # 使用 utf_8_sig 确保 Excel 打开映射表时不乱码
    mapping_df.to_csv(output_path, index=False, encoding='utf_8_sig')

    print(f"提取完成！已生成 {len(mapping_df)} 组唯一对应关系。")
    print(f"映射表已保存至: {output_path}")


# --- 参数设置 ---
file_name = '阳光优采交易订单2月.xlsx'  # 你的原始文件名
output_name = '采购部门企业对照表.csv'  # 导出的字典文件

extract_department_mapping_xlsx(file_name, output_name)