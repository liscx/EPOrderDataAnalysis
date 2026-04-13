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
            
            # --- 升级：智能关联推理识别 ---
            # 1. 首先构建本地 Sheet 内的身份关联字典 (为缺失行提供推断依据)
            # 过滤出所有已经有“采购企业”且非空的行
            valid_ref = df.dropna(subset=['采购企业', '采购人', '收货地址'])
            valid_ref = valid_ref[~valid_ref['采购企业'].astype(str).str.lower().isin(['nan', 'none', ''])]
            
            # 建立映射 (如果同一采购主体对应多个企业，保留最新的记录)
            person_to_ent = dict(zip(valid_ref['采购人'].astype(str).str.strip(), 
                                    valid_ref['采购企业'].astype(str).str.strip()))
            addr_to_ent = dict(zip(valid_ref['收货地址'].astype(str).str.strip(), 
                                    valid_ref['采购企业'].astype(str).str.strip()))

            def do_smart_fill(row):
                ent_raw = str(row.get('采购企业')).strip()
                dept_raw = str(row.get('采购部门')).strip()
                person_raw = str(row.get('采购人')).strip()
                addr_raw = str(row.get('收货地址')).strip()

                # 判断“采购企业”是否真正缺失
                is_missing = pd.isna(row.get('采购企业')) or ent_raw.lower() in ['nan', 'none', '']

                if is_missing:
                    # 轮次一：查外部字典 (基于部门)
                    if dept_raw in mapping_dict:
                        return mapping_dict[dept_raw]
                    
                    # 轮次二：看当前 Sheet 中的同一采购人
                    if person_raw in person_to_ent:
                        return person_to_ent[person_raw]
                    
                    # 轮次三：看当前数据中的相同收货地址
                    if addr_raw in addr_to_ent:
                        return addr_to_ent[addr_raw]
                    
                    # 轮次四：全部失败，使用明确标记占位，防止后续分析脚本中的 ffill() 跨订单乱填
                    return "采购企业不存在"
                
                return row.get('采购企业')

            print(f"      - 正在执行关联识别并补全 Sheet 数据...")
            
            # --- 审计记录：用于记录无法匹配的“孤儿订单” ---
            missing_audit_data = []

            def do_smart_fill(row):
                ent_raw = str(row.get('采购企业')).strip()
                dept_raw = str(row.get('采购部门')).strip()
                person_raw = str(row.get('采购人')).strip()
                addr_raw = str(row.get('收货地址')).strip()
                order_id = str(row.get('订单号')).strip()
                zone_name = str(row.get('专区名称')).strip()

                # 判断“采购企业”是否真正缺失
                is_missing = pd.isna(row.get('采购企业')) or ent_raw.lower() in ['nan', 'none', '']

                if is_missing:
                    # 轮次一：查外部字典 (基于部门)
                    if dept_raw in mapping_dict:
                        return mapping_dict[dept_raw]
                    
                    # 轮次二：看当前 Sheet 中的同一采购人
                    if person_raw in person_to_ent:
                        return person_to_ent[person_raw]
                    
                    # 轮次三：看当前数据中的相同收货地址
                    if addr_raw in addr_to_ent:
                        return addr_to_ent[addr_raw]
                    
                    # 轮次四：全部失败，记录审计信息并打标
                    is_real_order_start = not (order_id.lower() in ['nan', 'none', ''])
                    if is_real_order_start:
                        missing_audit_data.append(f"订单号: {order_id} | 专区: {zone_name} | 采购人: {person_raw} | 地址: {addr_raw}")
                        return "采购企业不存在"
                    
                    # 并非订单起始行，或者没有订单号的数据，保持原样（通常是空），不要乱填
                    return row.get('采购企业')
                
                return row.get('采购企业')

            df['采购企业'] = df.apply(do_smart_fill, axis=1)
            all_sheets[sn] = df

            # 写入“孤儿订单”审计日志
            if missing_audit_data:
                # 定位日志路径：放在与目标文件同级的目录下
                log_dir = os.path.dirname(os.path.abspath(target_file))
                log_path = os.path.join(log_dir, "企业名称缺失订单审计.txt")
                
                # 去重汇总
                unique_orphans = sorted(list(set(missing_audit_data)))
                with open(log_path, "w", encoding="utf-8") as f:
                    f.write("=" * 60 + "\n")
                    f.write(f"  企业名称缺失订单审计日志 (共计 {len(unique_orphans)} 笔异动)\n")
                    f.write("=" * 60 + "\n\n")
                    f.write("\n".join(unique_orphans))
                print(f"      📄 已生成异常审计日志: {log_path}")

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