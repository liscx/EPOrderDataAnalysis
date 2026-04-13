import pandas as pd
import os
import sys

def repair_enterprise_data(file_path):
    """
    根据相同采购人或相同收货地址，补全缺失的采购企业。
    
    逻辑：
    1. 如果 '采购企业' 缺失但 '订单号' 存在：
    2. 第一轮：查找相同 '采购人' 对应的 '采购企业'。
    3. 第二轮：查找相同 '收货地址' 对应的 '采购企业'。
    4. 如果都找不到，填充 '采购企业不存在'。
    5. 记录订单号到 '订单修复记录.txt'。
    """
    if not os.path.exists(file_path):
        print(f"❌ 修复失败：文件不存在 {file_path}")
        return

    print(f"-> [前置处理] 正在通过内部关联补全采购企业: {os.path.basename(file_path)}")
    
    try:
        # 读取所有 Sheet
        with pd.ExcelFile(file_path) as xls:
            all_sheets = {sn: pd.read_excel(xls, sheet_name=sn) for sn in xls.sheet_names}
    except Exception as e:
        print(f"❌ 读取 Excel 失败: {e}")
        return

    # 全局映射字典
    person_to_ent = {}
    addr_to_ent = {}
    
    # 辅助列，用于快速匹配
    for sn, df in all_sheets.items():
        # 清理列名空格
        df.columns = [str(c).strip() for c in df.columns]
        
        # 只要包含这三列，就提取有效映射
        required_cols = ['采购企业', '采购人', '收货地址']
        if all(col in df.columns for col in required_cols):
            # 过滤出有采购企业的行
            valid_df = df.dropna(subset=['采购企业', '采购人', '收货地址'])
            valid_df = valid_df[~valid_df['采购企业'].astype(str).str.strip().str.lower().isin(['nan', 'none', '', '采购企业不存在'])]
            
            for _, row in valid_df.iterrows():
                p = str(row['采购人']).strip()
                a = str(row['收货地址']).strip()
                e = str(row['采购企业']).strip()
                
                if p and p not in ['nan']:
                    person_to_ent[p] = e
                if a and a not in ['nan']:
                    addr_to_ent[a] = e

    # 执行修复
    repaired_orders = []
    not_found_orders = []
    
    for sn, df in all_sheets.items():
        if '订单号' not in df.columns:
            continue
            
        # 确保列都存在，不存在则补 dummy
        for col in ['采购企业', '采购人', '收货地址']:
            if col not in df.columns:
                df[col] = None

        def apply_repair(row):
            ent = str(row.get('采购企业')).strip()
            order_id = str(row.get('订单号')).strip()
            
            is_missing = pd.isna(row.get('采购企业')) or ent.lower() in ['nan', 'none', '']
            has_order = order_id.lower() not in ['nan', 'none', '']
            
            if is_missing and has_order:
                p = str(row.get('采购人')).strip()
                a = str(row.get('收货地址')).strip()
                
                # 第一轮：采购人
                if p in person_to_ent:
                    repaired_orders.append(f"Sheet: {sn} | 订单号: {order_id} | 匹配源: 采购人({p}) -> {person_to_ent[p]}")
                    return person_to_ent[p]
                
                # 第二轮：收货地址
                if a in addr_to_ent:
                    repaired_orders.append(f"Sheet: {sn} | 订单号: {order_id} | 匹配源: 地址({a}) -> {addr_to_ent[a]}")
                    return addr_to_ent[a]
                
                # 第三轮：失败
                not_found_orders.append(f"Sheet: {sn} | 订单号: {order_id} | 采购人: {p} | 地址: {a}")
                return "采购企业不存在"
            
            return row.get('采购企业')

        df['采购企业'] = df.apply(apply_repair, axis=1)
        all_sheets[sn] = df

    # 保存文件
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sn, df_data in all_sheets.items():
                df_data.to_excel(writer, sheet_name=sn, index=False)
        print("✅ 内部关联修复已写回原文件。")
    except Exception as e:
        print(f"❌ 写入修复数据失败: {e}")

    # 记录日志
    log_path = os.path.join(os.path.dirname(file_path), "订单修复记录.txt")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("=" * 60 + "\n")
        f.write(f"  采购企业缺失自动修复审计日志\n")
        f.write("=" * 60 + "\n\n")
        
        f.write(f"【修复成功的订单 (共 {len(repaired_orders)} 笔)】\n")
        if repaired_orders:
            f.write("\n".join(repaired_orders) + "\n")
        else:
            f.write("无\n")
            
        f.write(f"\n【仍匹配不到企业的订单 (共 {len(not_found_orders)} 笔)】\n")
        if not_found_orders:
            f.write("\n".join(not_found_orders) + "\n")
        else:
            f.write("无\n")
            
    print(f"-> 审计记录已保存至: {log_path}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        repair_enterprise_data(sys.argv[1])
    else:
        print("Usage: python enterprise_repair.py <excel_file_path>")
