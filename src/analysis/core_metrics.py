import pandas as pd
import yaml
import os
import traceback

def load_config():
    """自动路径识别的配置加载"""
    paths = ["config.yaml", "source/config.yaml", "../source/config.yaml"]
    for p in paths:
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                return yaml.safe_load(f)
    raise FileNotFoundError("无法找到 config.yaml，请检查路径。")

def run_enhanced_metrics():
    print("正在基于 [全量历史与清洗报表] 补齐核心话术模板遗落指标 (KV字典)...")
    try:
        config = load_config()
    except Exception as e:
        print(f"配置文件读取错误：{e}")
        return

    file_path = config.get('file_config', {}).get('output_file', '')
    if not file_path or not os.path.exists(file_path):
        print(f"数据文件不存在: {file_path}")
        return

    kv_data = {}
    
    try:
        xls = pd.ExcelFile(file_path)
        df_raw = xls.parse('清洗后数据')
        df_time_range = xls.parse('汇总_时间_区间')
        df_time_all = xls.parse('汇总_时间_全量')
        df_zone_all = xls.parse('汇总_专区_全量')
        df_buyer_sum = xls.parse('采购企业汇总表')
    except Exception as e:
        print(f"读取被依赖的历史汇总 Excel 失败: {e}")
        return
        
    # --- 1. 基础数据清洗与避坑 ---
    try:
        df_raw.columns = [str(c).strip() for c in df_raw.columns]
        buyer_col = '采购企业' if '采购企业' in df_raw.columns else '采购部门'
        if buyer_col not in df_raw.columns and len(df_raw.columns) > 5:
            buyer_col = df_raw.columns[5]
            
        fill_cols = ['订单号', '供应商', buyer_col, '专区名称']
        for col in fill_cols:
            if col in df_raw.columns:
                df_raw[col] = df_raw[col].ffill()
                
        if '订单状态' in df_raw.columns:
            df_raw['订单状态'] = df_raw['订单状态'].ffill()

        if '订单日期' in df_raw.columns:
            df_raw['订单日期'] = pd.to_datetime(df_raw['订单日期'], errors='coerce').ffill()
            
        df_raw['单价（元）'] = pd.to_numeric(df_raw.get('单价（元）', 0), errors='coerce').fillna(0)
        df_raw['数量'] = pd.to_numeric(df_raw.get('数量', 0), errors='coerce').fillna(0)
        df_raw['行金额_计算'] = df_raw['单价（元）'] * df_raw['数量']
    except Exception as e:
        print(f"数据清洗错误: {e}")

    try:
        conf_start = pd.to_datetime(config['analysis_period']['start_date'])
        conf_end = pd.to_datetime(config['analysis_period']['end_date'])
        target_month = conf_start.month
        last_month_str = (conf_start - pd.offsets.MonthBegin(1)).strftime('%Y年%m月')
    except Exception as e:
        print(f"时间处理错误: {e}")
        return

    # ===== 找回由于宽表隔离导致的全部缺失指标 ===== #
    # [话术1] 综合概况还原
    try:
        total_buyer = df_raw[buyer_col].nunique() if buyer_col in df_raw.columns else 0
        
        if '供应商' in df_raw.columns:
            # 全局清理供应商乱码，用于严格计数
            valid_sup_mask = ~df_raw['供应商'].astype(str).str.strip().isin(['nan', 'None', '', '汇总', '合计'])
            total_sup = df_raw[valid_sup_mask]['供应商'].nunique()
            ec_keys = ['得力', '齐心', '苏宁', '史泰博', '欧菲斯', '京东', '晨光', '震坤行']
            is_ec = df_raw['供应商'].apply(lambda x: any(k in str(x) for k in ec_keys))
            ec_count = df_raw[is_ec & valid_sup_mask]['供应商'].nunique()
        else:
            total_sup, ec_count = 0, 0

        all_order_count = df_raw['订单号'].nunique() if '订单号' in df_raw.columns else 0
        all_order_money = df_raw['行金额_计算'].sum()

        target_status = ['收货完成', '已完成发货', '已收货']
        if '订单状态' in df_raw.columns:
            df_done = df_raw[df_raw['订单状态'].isin(target_status)]
            done_count = df_done['订单号'].nunique()
            done_money = df_done['行金额_计算'].sum()
        else:
            done_count, done_money = 0, 0

        kv_data['{{采购人数量}}'] = str(total_buyer)
        kv_data['{{供应商数量}}'] = str(total_sup)
        kv_data['{{电商数量}}'] = str(ec_count)
        kv_data['{{本地供应商数量}}'] = str(total_sup - ec_count)
        kv_data['{{订单数量}}'] = str(all_order_count)
        kv_data['{{订单总额}}'] = f"{all_order_money:.2f}"
        kv_data['{{已完成订单数量}}'] = str(done_count)
        kv_data['{{已完成订单总额}}'] = f"{done_money:.2f}"
    except Exception as e:
        print(f"补全综合全量概况时出错: {e}")

    # [话术2] 累计历史大专区还原
    try:
        df_zone_summary = df_zone_all[df_zone_all['专区/时间'].astype(str).str.contains('小计', na=False)].copy()
        total_orders_all = df_zone_summary['订单数量'].sum()
        kv_data['{{产生订单专区数量}}'] = str(len(df_zone_summary[df_zone_summary['订单数量'] > 0]))

        # Bug修正：纠正旧代错位索引，按照订单实际数值真正抓出历史 TOP 2 的两大主力
        z_rank = df_zone_summary.sort_values('订单数量', ascending=False)
        if len(z_rank) > 0:
            top1 = z_rank.iloc[0]
            kv_data['{{主要专区1}}'] = str(top1['专区/时间']).replace(' 小计','')
            kv_data['{{主要专区1订单}}'] = str(int(top1['订单数量']))
            pct1 = (top1['订单数量'] / total_orders_all * 100) if total_orders_all > 0 else 0
            kv_data['{{主要专区1占比}}'] = f"{pct1:.2f}%"
        if len(z_rank) > 1:
            top2 = z_rank.iloc[1]
            kv_data['{{主要专区2}}'] = str(top2['专区/时间']).replace(' 小计','')
            kv_data['{{主要专区2订单}}'] = str(int(top2['订单数量']))
            pct2 = (top2['订单数量'] / total_orders_all * 100) if total_orders_all > 0 else 0
            kv_data['{{主要专区2占比}}'] = f"{pct2:.2f}%"
    except Exception as e:
        print(f"补全累计历史专区时出错: {e}")

    # [话术3] 本期主次专区表现与环比跨期还原
    try:
        kv_data['{{月份}}'] = str(target_month)
        
        tm_subtotal_mask = df_time_range['时间/专区'].astype(str).str.contains('小计', na=False)
        if not df_time_range[tm_subtotal_mask].empty:
            tm_row = df_time_range[tm_subtotal_mask].iloc[0]
            tm_money, tm_count = tm_row['交易金额(元)'], tm_row['订单数量']
        else:
            tm_money, tm_count = 0, 0

        lm_data = df_time_all[df_time_all['时间/专区'].astype(str).str.contains(f"{last_month_str} 小计", na=False)]
        lm_money = lm_data['交易金额(元)'].values[0] if not lm_data.empty else 0

        zone_rank_tm = df_time_range[~df_time_range['明细项'].astype(str).str.contains('---', na=False)]
        zone_rank_tm = zone_rank_tm.sort_values('交易金额(元)', ascending=False)

        kv_data['{{当月产生订单专区数量}}'] = str(len(zone_rank_tm))
        kv_data['{{当月订单数量}}'] = str(int(tm_count))
        kv_data['{{当月交易总额}}'] = f"{tm_money:.2f}"
        kv_data['{{上月交易总额}}'] = f"{lm_money:.2f}"
        kv_data['{{增长金额}}'] = f"{(tm_money - lm_money):.2f}"
        kv_data['{{环比增长率}}'] = f"{((tm_money - lm_money) / lm_money * 100):.2f}%" if lm_money > 0 else "0.00%"

        if len(zone_rank_tm) > 0:
            tr1 = zone_rank_tm.iloc[0]
            kv_data['{{主要贡献专区}}'] = str(tr1['明细项'])
            kv_data['{{主要贡献专区订单}}'] = str(int(tr1['订单数量']))
            kv_data['{{主要贡献专区总额}}'] = f"{tr1['交易金额(元)']:.2f}"
        if len(zone_rank_tm) > 1:
            tr2 = zone_rank_tm.iloc[1]
            kv_data['{{次要贡献专区}}'] = str(tr2['明细项'])
            kv_data['{{次要贡献专区订单}}'] = str(int(tr2['订单数量']))
            kv_data['{{次要贡献专区总额}}'] = f"{tr2['交易金额(元)']:.2f}"
    except Exception as e:
        print(f"补全当期跨期环比时出错: {e}")

    # [话术4] 基于外层原始报表的顶配指标复现
    try:
        if not df_buyer_sum.empty:
            df_buyer_sum.columns = [str(c).strip() for c in df_buyer_sum.columns]
            b_col_hist = '采购企业' if '采购企业' in df_buyer_sum.columns else df_buyer_sum.columns[1]
            
            b_ord = df_buyer_sum.sort_values('订单数量', ascending=False).iloc[0]
            kv_data['{{最高订单采购人}}'] = str(b_ord[b_col_hist])
            kv_data['{{最高订单数}}'] = str(int(b_ord['订单数量']))
            kv_data['{{最高订单总额}}'] = f"{b_ord['订单总额（元）']:.2f}"
            kv_data['{{最高订单专区}}'] = str(b_ord.get('专区名称', ''))
            
            b_mon = df_buyer_sum.sort_values('订单总额（元）', ascending=False).iloc[0]
            kv_data['{{最高金额采购人}}'] = str(b_mon[b_col_hist])
            kv_data['{{最高金额订单数}}'] = str(int(b_mon['订单数量']))
            kv_data['{{最高金额}}'] = f"{b_mon['订单总额（元）']:.2f}"   # 真正修补了那个错名
            kv_data['{{最高金额专区}}'] = str(b_mon.get('专区名称', ''))
            
            kv_data['{{活跃采购人数量}}'] = str(df_buyer_sum[b_col_hist].nunique())
    except Exception as e:
        print(f"补全历史最值排行榜时出错: {e}")

    # --- 剩余自己重写的更优字典逻辑合并补充 ---
    # 动态本地及各类特殊指标补充
    try:
        if '供应商' in df_raw.columns:
            valid_sup_mask = ~df_raw['供应商'].astype(str).str.strip().isin(['nan', 'None', '', '汇总', '合计'])
            ec_keys = ['得力', '齐心', '苏宁', '史泰博', '欧菲斯', '京东', '晨光', '震坤行']
            is_ec = df_raw['供应商'].apply(lambda x: any(k in str(x) for k in ec_keys))
            df_local = df_raw[~is_ec & valid_sup_mask]
            
            kv_data['{{本地供应商涉及数量}}'] = str(df_local['供应商'].nunique())
            kv_data['{{本地供应商金额}}'] = f"{df_local['行金额_计算'].sum():.2f}"
            
            sup_summary = df_local.groupby('供应商').agg(
                订单数量=pd.NamedAgg(column='订单号', aggfunc='nunique'),
                总金额=pd.NamedAgg(column='行金额_计算', aggfunc='sum')
            ).reset_index().sort_values(by='总金额', ascending=False)
            
            for i in range(len(sup_summary)):
                row = sup_summary.iloc[i]
                rank = i + 1
                kv_data[f'{{{{本地供应商{rank}}}}}'] = str(row['供应商'])
                kv_data[f'{{{{本地供应商{rank}订单}}}}'] = str(int(row['订单数量']))
                kv_data[f'{{{{本地供应商{rank}金额}}}}'] = f"{row['总金额']:.2f}"
    except: pass

    try:
        if '商品名称' in df_raw.columns:
            valid_items = df_raw[~df_raw['商品名称'].astype(str).isin(['nan', '', 'None', '汇总', '合计'])]
            kv_data['{{商品总数}}'] = str(valid_items['商品名称'].nunique())
            kv_data['{{商品品类}}'] = "办公耗材/电子产品"
    except: pass

    # --- 安全转码和出库 ---
    try:
        out_df = pd.DataFrame(list(kv_data.items()), columns=['占位符', '填充值'])
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            out_df.to_excel(writer, sheet_name="核心提取数据", index=False)
        print(">>> 核心话术指标 (满血完整版全映射) 修正补充完成！")
    except Exception as e:
        print(f"结果写入失败: {e}")

if __name__ == "__main__":
    run_enhanced_metrics()