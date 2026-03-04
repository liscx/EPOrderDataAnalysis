import pandas as pd
import yaml
import os


def load_config():
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def format_m(val):
    """金额格式化：千分位并保留两位小数"""
    try:
        v = str(val).replace(',', '')
        return f"{float(v):,.2f}"
    except:
        return "0.00"


def safe_split(val, expected_len):
    """安全拆分关联信息，防止索引越界"""
    if pd.isna(val):
        return ["0"] * expected_len
    parts = str(val).split('|')
    if len(parts) < expected_len:
        parts.extend(["0"] * (expected_len - len(parts)))
    return parts


def run_report_output():
    config = load_config()

    # 【核心逻辑】
    # 1. 从 output_file 获取 Excel 的完整路径
    excel_path = config['file_config']['output_file']
    # 2. 提取 output_file 所在的文件夹目录作为输出路径
    output_dir = os.path.dirname(excel_path)
    # 3. 构造 TXT 文件的完整保存路径
    txt_output_path = os.path.join(output_dir, "月报话术.txt")

    # 格式化截止日期
    end_date = pd.to_datetime(config['analysis_period']['end_date']).strftime('%Y年%m月%d日')

    if not os.path.exists(excel_path):
        print(f"错误：在路径 [{excel_path}] 下找不到输出 Excel 文件，请确认配置项 output_file 正确。")
        return

    # 读取核心数据 Sheet
    try:
        # 注意：这里直接读取 output_file 指向的文件，因为它包含了回填后的“核心话术数据”
        df_core = pd.read_excel(excel_path, sheet_name="核心话术数据")
        d = df_core.set_index('维度').to_dict('index')
    except Exception as e:
        print(f"读取 [核心话术数据] 失败。错误: {e}")
        return

    report_content = []
    report_content.append(f"{'=' * 20} 运营报告话术导出 ({end_date}) {'=' * 20}")

    # --- 1. 概况 ---
    if "话术1" in d:
        v = safe_split(d["话术1"]['数值'], 4)
        i = safe_split(d["话术1"]['关联信息'], 4)
        report_content.append(
            f"1.截至{end_date}，阳光优采产品累计注册采购人{v[0]}家、供应商{v[1]}家（电商{v[2]}家、本地供应商{v[3]}家），"
            f"已产生交易订单{i[0]}笔，订单总额{format_m(i[1])}元，其中已完成收货订单{i[2]}笔，订单总额{format_m(i[3])}元。")

    # --- 2. 累计专区 ---
    if "话术2" in d:
        count = d["话术2"]['数值']
        i = safe_split(d["话术2"]['关联信息'], 9)
        report_content.append(f"2.截至{end_date}，阳光优采产品已有{count}个专区产生交易订单，累计达{i[0]}笔。"
                              f"其中中国煤地电子商城{i[1]}笔，占比{i[2]}；新疆阳光采购平台{i[3]}笔，占比{i[4]}；"
                              f"大连市阳光采购服务平台{i[5]}笔，占比{i[6]}；邯郸市阳光优采平台{i[7]}笔，占比{i[8]}。")

    # --- 3. 本月表现 ---
    if "话术3" in d:
        v = safe_split(d["话术3"]['数值'], 3)
        i = safe_split(d["话术3"]['关联信息'], 8)
        tm_m, lm_m = float(i[0]), float(i[1])
        diff, rate = tm_m - lm_m, ((tm_m - lm_m) / lm_m * 100) if lm_m > 0 else 0
        report_content.append(f"3.{v[0]}月有{v[1]}个专区共产生交易订单{v[2]}笔，交易总金额达{format_m(tm_m)}元，"
                              f"较上月{format_m(lm_m)}元增长{format_m(diff)}元，环比增长{rate:.2f}%。其中，{i[2]}专区贡献主要交易体量，"
                              f"{v[0]}月完成{int(float(i[3]))}笔订单，订单总额{format_m(i[4])}元；"
                              f"{i[5]}同步发力，产生{int(float(i[6]))}笔订单，订单总额{format_m(i[7])}元。")

    # --- 4. 历史最值 ---
    if "话术4" in d:
        v_num = d["话术4"]['数值']
        i = safe_split(d["话术4"]['关联信息'], 8)
        report_content.append(f"4.截至{end_date}，交易活跃采购人{v_num}家。单采购人历史最高采购订单{int(float(i[0]))}笔"
                              f"（{i[1]}，订单总额{format_m(i[2])}元，来自{i[3]}专区），单采购人历史最高采购订单金额{format_m(i[4])}元"
                              f"（{i[5]}，订单数量{int(float(i[6]))}笔，来自{i[7]}专区）。")

    # --- 5. 新注册采购人 ---
    if "话术5" in d:
        month = safe_split(d.get("话术3", {}).get('数值', "0"), 1)[0]
        report_content.append(f"5.{month}月份新注册采购人{d['话术5']['数值']}家，来自于{d['话术5']['关联信息']}。")

    # --- 6. 交易供应商分类 ---
    if "话术6" in d:
        v = d["话术6"]['数值']
        i = safe_split(d["话术6"]['关联信息'], 2)
        report_content.append(f"6.截至{end_date}，产生交易订单供应商{v}家，其中{i[0]}家为电商企业，{i[1]}家为本地供应商。")

    # --- 7. 本地表现 ---
    if "话术7" in d:
        month = safe_split(d.get("话术3", {}).get('数值', "0"), 1)[0]
        count = d["话术7"]['数值']
        i = safe_split(d["话术7"]['关联信息'], 4)
        report_content.append(
            f"7.{month}月份本地供应商共涉及{count}家供应商，订单金额{format_m(i[0])}元（其中 {i[1]}，{int(float(i[2]))}笔订单，订单金额{format_m(i[3])}元）。")

    # --- 8. 新注册供应商 ---
    if "话术8" in d:
        month = safe_split(d.get("话术3", {}).get('数值', "0"), 1)[0]
        report_content.append(f"8.{month}月新注册供应商{d['话术8']['数值']}家，来自于{d['话术8']['关联信息']}。")

    # --- 9. 商品概括 ---
    if "话术9" in d:
        month = safe_split(d.get("话术3", {}).get('数值', "0"), 1)[0]
        report_content.append(f"9.{month}月产生交易商品{d['话术9']['数值']}件，涵盖{d['话术9']['关联信息']}等多个品类。")

    # 执行保存
    try:
        with open(txt_output_path, "w", encoding="utf-8") as f:
            f.write("\n\n".join(report_content))
        print(f"--- 话术提取成功 ---")
        print(f"输出位置：{txt_output_path}")
    except Exception as e:
        print(f"写入 TXT 失败: {e}")


if __name__ == "__main__":
    run_report_output()