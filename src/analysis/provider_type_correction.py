import pandas as pd
import os
from ..core import load_config
# 注意：这里我们从 constants 导入最新的供应商名单
from ..core.constants import LOCAL_PROVIDER_NAME, EC_PROVIDER_NAME

def run_provider_type_correction():
    # 1. 加载基础配置
    try:
        config = load_config()
    except Exception as e:
        print(f"错误：加载配置失败: {e}")
        return

    file_path = config['file_config']['output_file']
    log_file = os.path.join(os.path.dirname(file_path), "供应商类型识别日志.txt")

    if not os.path.exists(file_path):
        print(f"提示：找不到文件 {file_path}，跳过类型自动修正。")
        return

    print("--- 正在执行：供应商类型自动对齐与纠偏 ---")
    
    # 2. 读取当前已存在的所有 Sheets (保证不丢失数据)
    with pd.ExcelFile(file_path) as xls:
        all_sheets = {sn: pd.read_excel(xls, sheet_name=sn) for sn in xls.sheet_names}

    audit_logs = []
    correction_count = 0

    # 3. 核心修正函数
    def correct_type(row):
        nonlocal correction_count
        p_name = str(row.get('供应商')).strip()
        p_type = str(row.get('供应商类型')).strip()

        # 第一级判定：是否在“本地供应商”白名单中
        if p_name in LOCAL_PROVIDER_NAME:
            target = "本地供应商"
            if p_type != target:
                audit_logs.append(f"【修正】供应商 [{p_name}] 类型从 [{p_type}] 修正为 [{target}]")
                correction_count += 1
            return target
        
        # 第二级判定：是否在“电商供应商”白名单中
        elif p_name in EC_PROVIDER_NAME:
            target = "电商供应商"
            if p_type != target:
                audit_logs.append(f"【修正】供应商 [{p_name}] 类型从 [{p_type}] 修正为 [{target}]")
                correction_count += 1
            return target
        
        # 第三级：不在白名单中，记录日志并保持原样
        else:
            if p_name not in ['nan', 'None', '']:
                # 仅记录未入名单的供应商详情
                audit_logs.append(f"【记录】供应商 [{p_name}] 不在预设名单中，维持原始类型: [{p_type}]")
            return row.get('供应商类型')

    # 4. 执行修正 (仅针对核心业务 Sheet)
    target_sheets = ['清洗后数据', '指定区间数据']
    for sn in target_sheets:
        if sn in all_sheets:
            print(f"   -> 正在扫描 [{sn}]...")
            all_sheets[sn]['供应商类型'] = all_sheets[sn].apply(correct_type, axis=1)

    # 5. 写回文件并保存日志
    try:
        # 去重日志
        unique_logs = sorted(list(set(audit_logs)))
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sn, df_data in all_sheets.items():
                df_data.to_excel(writer, sheet_name=sn, index=False)
        
        with open(log_file, "w", encoding="utf-8") as f:
            f.write("=" * 60 + "\n")
            f.write(f"  供应商类型审计与自动纠偏日志 (产生的修正: {correction_count} 处)\n")
            f.write("=" * 60 + "\n\n")
            f.write("\n".join(unique_logs))
            
        print(f"✅ 类型自动修正完成！累计修正 {correction_count} 处。")
        print(f"📄 审计详情已记录至: {log_file}")
    except Exception as e:
        print(f"❌ 写入 Excel 失败，请检查文件是否被占用。详情: {e}")


if __name__ == "__main__":
    run_provider_type_correction()
