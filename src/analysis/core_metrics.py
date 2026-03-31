import pandas as pd
import yaml
import os
import traceback

from ..core import load_config

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

    # ==============================================================
    # 以下为指标占位符的提取与计算模块
    # ==============================================================
    
    # 【1】截止日期
    # 从 config.yaml 的 analysis_period 配置中读取 end_date
    try:
        if 'end_date' in config.get('analysis_period', {}):
            d_str = config['analysis_period']['end_date']
            d_dt = pd.to_datetime(d_str)
            kv_data['{{截止日期}}'] = d_dt.strftime('%Y年%m月%d日')
    except Exception as e:
        print(f"提取 {{截止日期}} 失败: {e}")

    # 【2】采购维度：采购人数量 (剔除空值、合并与无意义项，按“采购企业”去重)
    try:
        if '采购企业' in df_raw.columns:
            valid_buyers = df_raw['采购企业'].astype(str).str.strip()
            total_buyer = valid_buyers[~valid_buyers.isin(['nan', 'None', '', '汇总', '合计'])].nunique()
            kv_data['{{采购人数量}}'] = str(total_buyer)
        else:
            kv_data['{{采购人数量}}'] = "0"
    except Exception as e:
        print(f"提取 {{采购人数量}} 失败: {e}")

    # 【3】供应商与电商维度：供应商数量、电商数量、本地供应商数量
    # 基于【清洗后数据】，严格按照“供应商”列去重统计
    try:
        if '供应商' in df_raw.columns:
            valid_sups = df_raw['供应商'].astype(str).str.strip()
            
            # 获取有效的供应商（非空且非汇总行）作为后续统计的基础范围
            valid_sup_mask = ~valid_sups.isin(['nan', 'None', '', '汇总', '合计'])
            
            # {{供应商数量}} 和 {{供应商总数}}：针对所有有效记录去重
            total_sup = valid_sups[valid_sup_mask].nunique()
            kv_data['{{供应商数量}}'] = str(total_sup)
            kv_data['{{供应商总数}}'] = str(total_sup)
            
            # 根据“供应商类型”列判断是否为电商（不再使用硬编码关键词库）
            if '供应商类型' in df_raw.columns:
                is_ec = df_raw['供应商类型'].astype(str).str.contains('电商', na=False)
            else:
                # 如果缺少该列，则默认都不是电商，避免报错
                is_ec = pd.Series([False] * len(df_raw))
            
            # {{电商数量}}：基于“供应商类型为电商”进行去重
            ec_count = valid_sups[is_ec & valid_sup_mask].nunique()
            kv_data['{{电商数量}}'] = str(ec_count)
            
            # {{本地供应商数量}}：基于“供应商类型非电商”进行去重
            local_count = valid_sups[(~is_ec) & valid_sup_mask].nunique()
            kv_data['{{本地供应商数量}}'] = str(local_count)
        else:
            kv_data['{{供应商数量}}'] = "0"
            kv_data['{{供应商总数}}'] = "0"
            kv_data['{{电商数量}}'] = "0"
            kv_data['{{本地供应商数量}}'] = "0"
    except Exception as e:
        print(f"提取 供应商相关数量 失败: {e}")

    # 【4】整体业务规模：订单数量、订单总额、已完成订单数量及总额
    # 按照指示，目前全量清洗后的数据（df_raw）即可被视作所有完成（合规未废除）数据的全量
    try:
        # {{订单数量}} 和 {{订单总额}}
        all_order_count = df_raw['订单号'].nunique() if '订单号' in df_raw.columns else 0
        all_order_money = df_raw['行金额_计算'].sum() if '行金额_计算' in df_raw.columns else 0
        kv_data['{{订单数量}}'] = str(all_order_count)
        kv_data['{{订单总额}}'] = f"{all_order_money:.2f}"

        # {{已完成订单数量}} 和 {{已完成订单总额}}
        # 读取“清洗后数据”中订单状态为“收货完成”的数据作为统计基础
        if '订单状态' in df_raw.columns:
            df_finished = df_raw[df_raw['订单状态'].astype(str).str.strip().str.contains('收货完成', na=False)]
            finished_count = df_finished['订单号'].nunique() if '订单号' in df_finished.columns else 0
            finished_money = df_finished['行金额_计算'].sum() if '行金额_计算' in df_finished.columns else 0
            kv_data['{{已完成订单数量}}'] = str(finished_count)
            kv_data['{{已完成订单总额}}'] = f"{finished_money:.2f}"
        else:
            kv_data['{{已完成订单数量}}'] = "0"
            kv_data['{{已完成订单总额}}'] = "0.00"
    except Exception as e:
        print(f"提取 订单整体指标 失败: {e}")

    # 【5】历史专区维度的贡献表现：覆盖区域数量以及最吸金专区
    # 基于【汇总_专区_全量】表的历史累加值分析
    try:
        # 抓取出所有带有“小计”标志的记录（即各个历史专区的总合）
        df_zone_summary = df_zone_all[df_zone_all['专区/时间'].astype(str).str.contains('小计', na=False)].copy()
        total_orders_all = df_zone_summary['订单数量'].sum()
        
        # {{产生订单专区数量}}：只要历史里产生出非 0 订单量的专区数量
        kv_data['{{产生订单专区数量}}'] = str(len(df_zone_summary[df_zone_summary['订单数量'] > 0]))

        # 对提取出来的小计按订单量排序，找出历史长河中的【主要专区1】和【主要专区2】
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
        print(f"提取 历史主次专区 失败: {e}")

    # 【6】区间时间维度的跨期比率（核心月度总结数据）
    # 当月主要指代执行配置指定的月份目标
    try:
        # {{月份}}
        kv_data['{{月份}}'] = str(target_month)
        
        # 定位当月的汇总小结：在【汇总_时间_区间】获取当月累计总金额/订单
        tm_subtotal_mask = df_time_range['时间/专区'].astype(str).str.contains('小计', na=False)
        if not df_time_range[tm_subtotal_mask].empty:
            tm_row = df_time_range[tm_subtotal_mask].iloc[0]
            tm_money, tm_count = tm_row['交易金额(元)'], tm_row['订单数量']
        else:
            tm_money, tm_count = 0, 0

        # 定位上月的汇总小结：作为动态基数（基于配置时间的上个月推导，并在历史总表里寻找）
        lm_data = df_time_all[df_time_all['时间/专区'].astype(str).str.contains(f"{last_month_str} 小计", na=False)]
        lm_money = lm_data['交易金额(元)'].values[0] if not lm_data.empty else 0

        # 获取当月各大专区的交易明细，并按金额降序排列（去除整体汇总的---项）
        zone_rank_tm = df_time_range[~df_time_range['明细项'].astype(str).str.contains('---', na=False)]
        zone_rank_tm = zone_rank_tm.sort_values('交易金额(元)', ascending=False)

        # 录入各项时间属性指标
        kv_data['{{当月产生订单专区数量}}'] = str(len(zone_rank_tm))
        kv_data['{{当月订单数量}}'] = str(int(tm_count))
        kv_data['{{当月交易总额}}'] = f"{tm_money:.2f}"
        kv_data['{{上月交易总额}}'] = f"{lm_money:.2f}"
        
        # 录入环比动态数据
        kv_data['{{增长金额}}'] = f"{(tm_money - lm_money):.2f}"
        kv_data['{{环比增长率}}'] = f"{((tm_money - lm_money) / lm_money * 100):.2f}%" if lm_money > 0 else "0.00%"

        # 分别将当月成交最活跃的俩大“主要贡献专区”作为业务骨架点明
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
        print(f"提取 当期跨期环比及排名表现 失败: {e}")

    # 【7】单一极值抓取：历史金牌榜（采购大户）
    # 针对【采购企业汇总表】取按次数最多和按金额最高的两名大头
    try:
        if not df_buyer_sum.empty:
            # 标准化表头以找到买方字段
            df_buyer_sum.columns = [str(c).strip() for c in df_buyer_sum.columns]
            b_col_hist = '采购企业' if '采购企业' in df_buyer_sum.columns else df_buyer_sum.columns[1]
            
            # 最高次数大户
            b_ord = df_buyer_sum.sort_values('订单数量', ascending=False).iloc[0]
            kv_data['{{最高订单采购人}}'] = str(b_ord[b_col_hist])
            kv_data['{{最高订单数}}'] = str(int(b_ord['订单数量']))
            kv_data['{{最高订单总额}}'] = f"{b_ord['订单总额（元）']:.2f}"
            kv_data['{{最高订单专区}}'] = str(b_ord.get('专区名称', ''))
            
            # 最高金额大户
            b_mon = df_buyer_sum.sort_values('订单总额（元）', ascending=False).iloc[0]
            kv_data['{{最高金额采购人}}'] = str(b_mon[b_col_hist])
            kv_data['{{最高金额订单数}}'] = str(int(b_mon['订单数量']))
            kv_data['{{最高金额}}'] = f"{b_mon['订单总额（元）']:.2f}"   
            kv_data['{{最高金额专区}}'] = str(b_mon.get('专区名称', ''))
            
            # {{活跃采购人数量}} (即在历史中有单量的唯一机构数)
            kv_data['{{活跃采购人数量}}'] = str(df_buyer_sum[b_col_hist].nunique())
    except Exception as e:
        print(f"提取 单极金牌排行榜 失败: {e}")

    # 【8】本地供应商销售排名专题：特定分析要求
    # 仍立足【清洗后数据】，独立剔除全国主流平台巨头后分析本地实体的分布
    try:
        if '供应商' in df_raw.columns:
            valid_sup_mask = ~df_raw['供应商'].astype(str).str.strip().isin(['nan', 'None', '', '汇总', '合计'])
            
            # 根据“供应商类型”列判断是否为电商
            if '供应商类型' in df_raw.columns:
                is_ec = df_raw['供应商类型'].astype(str).str.contains('电商', na=False)
            else:
                is_ec = pd.Series([False] * len(df_raw))

            df_local = df_raw[~is_ec & valid_sup_mask]
            
            # {{本地供应商涉及数量}} 与 {{本地供应商金额}} 
            kv_data['{{本地供应商涉及数量}}'] = str(df_local['供应商'].nunique())
            kv_data['{{本地供应商金额}}'] = f"{df_local['行金额_计算'].sum():.2f}"
            
            # 聚合本地商户业绩，按金额倒序
            sup_summary = df_local.groupby('供应商').agg(
                订单数量=pd.NamedAgg(column='订单号', aggfunc='nunique'),
                总金额=pd.NamedAgg(column='行金额_计算', aggfunc='sum')
            ).reset_index().sort_values(by='总金额', ascending=False)
            
            # 为前排打入 {{本地供应商1}} 等序号标记，模板使用循环遍历进行展示
            for i in range(len(sup_summary)):
                row = sup_summary.iloc[i]
                rank = i + 1
                kv_data[f'{{{{本地供应商{rank}}}}}'] = str(row['供应商'])
                kv_data[f'{{{{本地供应商{rank}订单}}}}'] = str(int(row['订单数量']))
                kv_data[f'{{{{本地供应商{rank}金额}}}}'] = f"{row['总金额']:.2f}"
    except Exception as e:
        print(f"提取 本地专题供应商指标 失败: {e}")

    # 【9】商品粗分类统计
    try:
        if '商品名称' in df_raw.columns:
            # 清除汇总行、空值，取得唯一的商品销售条目总量
            valid_items = df_raw[~df_raw['商品名称'].astype(str).isin(['nan', '', 'None', '汇总', '合计'])]
            kv_data['{{商品总数}}'] = str(valid_items['商品名称'].nunique())
            kv_data['{{商品品类}}'] = "办公耗材/电子产品"
    except Exception as e:
        print(f"提取 粗分类统统 失败: {e}")

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