import pandas as pd
import os
import sys

# 尝试引入常量
try:
    # 这里的路径处理需要兼容从 workflow.py 调用
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    src_path = os.path.join(project_root, "src")
    if src_path not in sys.path:
        sys.path.insert(0, src_path)
    
    from core.constants import LOCAL_PROVIDER_NAME, EC_PROVIDER_NAME
except ImportError:
    # 备用方案（如果 sys.path 没设好）
    LOCAL_PROVIDER_NAME = []
    EC_PROVIDER_NAME = []

def repair_provider_type(file_path):
    """
    根据 constants.py 中的配置，修正供应商类型。
    
    逻辑：
    1. 遍历所有 Sheet 的 '供应商' 列。
    2. 如果在 LOCAL_PROVIDER_NAME 中，修正 '供应商类型' 为 '本地'。
    3. 如果在 EC_PROVIDER_NAME 中，修正 '供应商类型' 为 '电商'。
    4. 如果都不在，保持原有的 '供应商类型'，并将供应商名称记录到日志。
    """
    if not os.path.exists(file_path):
        print(f"❌ 修正失败：文件不存在 {file_path}")
        return

    print(f"-> [前置处理] 正在校验并修正供应商类型: {os.path.basename(file_path)}")
    
    try:
        xls = pd.ExcelFile(file_path)
        all_sheets = {sn: xls.parse(sn) for sn in xls.sheet_names}
    except Exception as e:
        print(f"❌ 读取 Excel 失败: {e}")
        return

    # 为了更好的匹配，预处理常量列表
    local_set = set(str(name).strip() for name in LOCAL_PROVIDER_NAME)
    ec_set = set(str(name).strip() for name in EC_PROVIDER_NAME)
    
    # 使用 mutable 列表存储统计信息，规避部分 linter 对闭包变量变异的误读
    repair_stats = {
        'count': 0,
        'unknown': {} # {供应商名称: {原有类型1, 原有类型2}}
    }

    for sn, df in all_sheets.items():
        # 清理列名
        df.columns = [str(c).strip() for c in df.columns]
        
        if '供应商' not in df.columns:
            continue
            
        if '供应商类型' not in df.columns:
            df['供应商类型'] = None

        def check_row_type(row):
            provider = str(row.get('供应商', '')).strip()
            current_type_raw = row.get('供应商类型')
            current_type = str(current_type_raw).strip() if pd.notna(current_type_raw) else "未知"
            
            if not provider or provider.lower() in ['nan', 'none', '']:
                return current_type_raw

            if provider in local_set:
                if current_type != '本地':
                    repair_stats['count'] += 1
                return '本地'
            elif provider in ec_set:
                if current_type != '电商':
                    repair_stats['count'] += 1
                return '电商'
            else:
                # 不在配置中，记录原有分类名
                if provider not in repair_stats['unknown']:
                    repair_stats['unknown'][provider] = set()
                repair_stats['unknown'][provider].add(current_type)
                return current_type_raw

        df['供应商类型'] = df.apply(check_row_type, axis=1)
        all_sheets[sn] = df

    # 保存
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sn, df_data in all_sheets.items():
                df_data.to_excel(writer, sheet_name=sn, index=False)
        print(f"✅ 供应商类型修正完成 (共修正 {repair_stats['count']} 行数据)。")
    except Exception as e:
        print(f"❌ 写入修正数据失败: {e}")

    # 记录未知供应商到日志
    log_path = os.path.join(os.path.dirname(file_path), "供应商类型异常.txt")
    
    if repair_stats['unknown']:
        unknown_lines = []
        for p, type_set in sorted(repair_stats['unknown'].items()):
            # 清理类型集合中的重复或无效值
            clean_types = [t for t in type_set if t not in ['nan', 'None', '']]
            type_desc = " / ".join(clean_types) if clean_types else "无"
            unknown_lines.append(f"{p.ljust(30)} | 已保持原有分类： {type_desc}")

        with open(log_path, "w", encoding="utf-8") as f:
            f.write("=" * 70 + "\n")
            f.write(f"  未定义供应商异常情况记录 (已保持原有分类)\n")
            f.write("=" * 70 + "\n\n")
            f.write("\n".join(unknown_lines))
        print(f"-> 发现 {len(repair_stats['unknown'])} 个未定义供应商，已记录详细信息至: {log_path}")
    else:
        if os.path.exists(log_path):
            try:
                os.remove(log_path)
            except:
                pass

if __name__ == "__main__":
    if len(sys.argv) > 1:
        repair_provider_type(sys.argv[1])
    else:
        print("Usage: python provider_type_repair.py <excel_file_path>")
