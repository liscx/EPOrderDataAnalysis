import pandas as pd
import os
import sys

# 尝试引入常量
try:
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    src_path = os.path.join(project_root, "src")
    if src_path not in sys.path:
        sys.path.insert(0, src_path)
    
    from core.constants import PROVIDER_ALIAS
except ImportError:
    PROVIDER_ALIAS = {}

def repair_provider_alias(file_path):
    """
    根据 constants.py 中的 PROVIDER_ALIAS，统一供应商名称。
    
    逻辑：
    1. 遍历所有 Sheet 的 '供应商' 列。
    2. 如果命中了别名映射，将其替换为标准名称。
    """
    if not os.path.exists(file_path):
        print(f"❌ 别名修复失败：文件不存在 {file_path}")
        return

    if not PROVIDER_ALIAS:
        print("-> [跳过] constants.py 中未配置供应商别名映射。")
        return

    print(f"-> [前置处理] 正在统一供应商名称 (Alias修复): {os.path.basename(file_path)}")
    
    try:
        xls = pd.ExcelFile(file_path)
        all_sheets = {sn: xls.parse(sn) for sn in xls.sheet_names}
    except Exception as e:
        print(f"❌ 读取 Excel 失败: {e}")
        return

    stats = {
        'total_fixed': 0,
        'details': []
    }

    for sn, df in all_sheets.items():
        # 清理列名
        df.columns = [str(c).strip() for c in df.columns]
        
        if '供应商' not in df.columns:
            continue

        def apply_alias(val):
            if pd.isna(val):
                return val
            
            s_val = str(val).strip()
            if s_val in PROVIDER_ALIAS:
                standard_name = PROVIDER_ALIAS[s_val]
                stats['total_fixed'] += 1
                stats['details'].append(f"Sheet: {sn} | {s_val} -> {standard_name}")
                return standard_name
            return val

        df['供应商'] = df['供应商'].apply(apply_alias)
        all_sheets[sn] = df

    if stats['total_fixed'] > 0:
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                for sn, df_data in all_sheets.items():
                    df_data.to_excel(writer, sheet_name=sn, index=False)
            print(f"✅ 供应商名称统一完成 (共修正 {stats['total_fixed']} 处)。")
            
            # 记录日志
            log_path = os.path.join(os.path.dirname(file_path), "供应商名称统一记录.txt")
            with open(log_path, "w", encoding="utf-8") as f:
                f.write("=" * 60 + "\n")
                f.write(f"  供应商名称别名统一审计日志\n")
                f.write("=" * 60 + "\n\n")
                f.write("\n".join(stats['details']))
            print(f"-> 详细修正记录已保存至: {log_path}")
            
        except Exception as e:
            print(f"❌ 写入修正数据失败: {e}")
    else:
        print("-> 无需修正供应商别名。")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        repair_provider_alias(sys.argv[1])
    else:
        print("Usage: python provider_alias_repair.py <excel_file_path>")
