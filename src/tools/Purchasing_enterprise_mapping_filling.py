import pandas as pd
import os


def fill_in_place(target_file, mapping_file):
    # 1. 加载映射表
    try:
        mapping_df = pd.read_csv(mapping_file, encoding='utf_8_sig')
    except:
        try:
            mapping_df = pd.read_csv(mapping_file, encoding='gbk')
        except:
            print("❌ 错误：映射表读取失败，请检查 CSV 编码。")
            return

    mapping_dict = dict(zip(mapping_df['采购部门'].astype(str).str.strip(),
                            mapping_df['采购企业'].astype(str).str.strip()))

    # 2. 读取目标原文件中所有的 Sheets
    print(f"正在读取文件: {target_file}")
    try:
        with pd.ExcelFile(target_file) as xls:
            all_sheets = {sn: pd.read_excel(xls, sheet_name=sn) for sn in xls.sheet_names}
    except Exception as e:
        print(f"❌ 读取 Excel 失败: {e}")
        return

    # 3. 执行填充 (对所有含有对应列的 Sheet 执行填充)
    print(f"   -> 正在扫描并补全 Sheet 数据...")
    found_any = False
    for sn, df in all_sheets.items():
        # 只要该 Sheet 包含 '采购部门'，就尝试回填
        if '采购部门' in df.columns:
            print(f"      - 正在处理 Sheet: [{sn}]")
            found_any = True
            
            def do_fill(row):
                dept = str(row['采购部门']).strip()
                # 如果当前企业名称为空(NaN)或者属于无效字符串，则从字典查找
                if pd.isna(row.get('采购企业')) or str(row.get('采购企业')).strip() == '' or str(row.get('采购企业')) == 'nan':
                    return mapping_dict.get(dept, row.get('采购企业'))
                return row.get('采购企业')
            
            df['采购企业'] = df.apply(do_fill, axis=1)
            all_sheets[sn] = df
    
    if not found_any:
        print("   ⚠️ 警告：在所有 Sheet 中均未找到 '采购部门' 列，无法执行回填。")

    # 4. 写回原文件 (保留所有原始 Sheets)
    try:
        with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
            for sn, df_data in all_sheets.items():
                df_data.to_excel(writer, sheet_name=sn, index=False)
        print("✅ --- 填充成功！所有 Sheet 已安全更新并回写 ---")
    except PermissionError:
        print("❌ 保存失败：请先关闭 Excel 程序后再运行脚本！")
    except Exception as e:
        print(f"❌ 写入阶段发生未知错误: {e}")


if __name__ == "__main__":
    # --- 设置参数 (仅用于本地调试) ---
    target_file = '阳光优采交易订单.xlsx'  # 你的原文件路径
    mapping_file = '采购部门企业对照表.csv'  # 映射字典路径
    fill_in_place(target_file, mapping_file)