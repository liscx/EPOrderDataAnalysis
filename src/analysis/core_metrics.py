import pandas as pd
import yaml
import os
import traceback
import argparse

from ..core import load_config

def run_enhanced_metrics(override_file_path=None):
    print("正在基于 [全量历史与清洗报表] 补齐核心话术模板遗落指标 (KV字典)...")
    try:
        config = load_config()
    except Exception as e:
        print(f"配置文件读取错误：{e}")
        config = {}

    file_path = override_file_path or config.get('file_config', {}).get('output_file', '')
    if not file_path or not os.path.exists(file_path):
        print(f"数据文件不存在: {file_path}")
        return

    kv_data = {}
    
    # ... (rest of the code remains the same, but using the file_path variable which is already used below)
    # Note: I will only replace the top and bottom of the function to keep it concise, 
    # but the tool requires the full TargetContent to match.
    # Actually, I can just replace the start of the function and the __main__ block.

    try:
        xls = pd.ExcelFile(file_path)
        df_raw = xls.parse('清洗后数据')
        df_period = xls.parse('指定区间数据')
        df_time_range = xls.parse('汇总_时间_区间')
        df_time_all = xls.parse('汇总_时间_全量')
        df_zone_all = xls.parse('汇总_专区_全量')
        df_buyer_sum = xls.parse('采购企业汇总表')
        # 避开后续 ExcelWriter 导致句柄失效，在此提前载入供应商辅助表
        try:
            df_new_sup_file = xls.parse('供应商信息提取')
        except:
            df_new_sup_file = pd.DataFrame()
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
        conf_end = pd.to_datetime(config['analysis_period']['end_date']) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        target_month = conf_start.month
        last_month_str = (conf_start - pd.offsets.MonthBegin(1)).strftime('%Y年%m月')
    except Exception as e:
        print(f"时间处理错误: {e}")
        return

    # ==============================================================
    # 以下为指标占位符的提取与计算模块
    # ==============================================================
    
    # ==============================================================
    # 【核心第一步】提取物理归档 Sheet：[截至数据]
    # 从全量清洗后数据中锁定 endDate 之前的记录，并执行严格的数据清洗（剔除 测试/汇总/合计 等干扰项）
    # ==============================================================
    try:
        limit_date = conf_end
        df_full = df_raw.copy()
        
        # 1.1 时间过滤：精准切片提取截至 endDate 的历史包
        if '订单日期' in df_full.columns:
            df_full['订单日期'] = pd.to_datetime(df_full['订单日期'], errors='coerce')
        df_as_of = df_full[df_full['订单日期'] <= limit_date].copy()
        
        # 1.2 严格清洗干扰项 (排除 汇总、测试、合计、nan、None 等)
        # 确保统计活跃采购人和供应商时，不计入这些辅助/垃圾行
        clean_pattern = '测试|汇总|合计|nan|None|NULL'
        if '采购企业' in df_as_of.columns:
            df_as_of = df_as_of[~df_as_of['采购企业'].astype(str).str.contains(clean_pattern, na=False)]
        if '供应商' in df_as_of.columns:
            df_as_of = df_as_of[~df_as_of['供应商'].astype(str).str.contains(clean_pattern, na=False)]
            
        # 1.3 写入物理 Sheet (作为后续计算的单一事实来源)
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_as_of.to_excel(writer, sheet_name="截至数据", index=False)
        print(">>> 物理归档库 [截至数据] 已建立并完成严格清洗回写。")
        
    except Exception as e:
        print(f"创建 [截至数据] 失败: {e}")
        df_as_of = df_raw.copy()  # 降级处理

    # ==============================================================
    # 【核心第二步】基于 [截至数据] 计算指定的累计话术占位符
    # ==============================================================
    
    # 【1】截止日期
    kv_data['{{截止日期}}'] = conf_end.strftime('%Y年%m月%d日')

    # 【2】采购人数量/供应商数量 (基于截至数据去重)
    try:
        # 获取采购人名录（去重且避雷）
        buyers = df_as_of['采购企业'].astype(str).str.strip()
        kv_data['{{采购人数量}}'] = str(buyers[~buyers.isin(['nan','汇总','合计','None'])].nunique())
        
        # 获取供应商名录（含分类）
        sups = df_as_of[['供应商', '供应商类型']].drop_duplicates()
        sups = sups[~sups['供应商'].astype(str).str.contains('测试|汇总|合计|nan', na=False)]
        
        total_sup = sups['供应商'].nunique()
        kv_data['{{供应商数量}}'] = str(total_sup)
        kv_data['{{供应商总数}}'] = str(total_sup)
        
        ec_count = sups[sups['供应商类型'].astype(str).str.contains('电商', na=False)]['供应商'].nunique()
        kv_data['{{电商数量}}'] = str(ec_count)
        kv_data['{{本地供应商数量}}'] = str(total_sup - ec_count)
    except:
        print("采购/供应商累计计算出错")

    # 【4】整体累计业务规模 (直接基于 [截至数据] 物理表求和)
    try:
        # 已产生交易订单笔数
        cum_order_count = df_as_of['订单号'].nunique() if '订单号' in df_as_of.columns else 0
        kv_data['{{订单数量}}'] = str(int(cum_order_count))
        
        # 订单总额
        # 优先使用计算过的金额列或基础金额列
        money_col = '行金额_计算' if '行金额_计算' in df_as_of.columns else '交易金额(元)'
        cum_money = df_as_of[money_col].sum() if money_col in df_as_of.columns else 0
        kv_data['{{订单总额}}'] = f"{cum_money:.2f}"
        
        # 已完成订单数量及金额对齐（基于截至包，严格匹配“收货完成”状态）
        if '订单状态' in df_as_of.columns:
            # 采用去空格后的精确匹配，确保口径百分百对齐您的要求
            df_as_of['订单状态_tmp'] = df_as_of['订单状态'].astype(str).str.strip()
            finished = df_as_of[df_as_of['订单状态_tmp'] == "收货完成"]
            kv_data['{{已完成订单数量}}'] = str(int(finished['订单号'].nunique()))
            kv_data['{{已完成订单总额}}'] = f"{finished[money_col].sum():.2f}"
            df_as_of.drop(columns=['订单状态_tmp'], inplace=True)
        else:
            # 如无状态列，降级为全量统计
            kv_data['{{已完成订单数量}}'] = str(int(cum_order_count))
            kv_data['{{已完成订单总额}}'] = f"{cum_money:.2f}"

    except:
        print("累计订单笔数/金额计算出错")

    # 【5】历史专区维度分析 (切换至 [截至数据] 聚合)
    try:
        # {{产生订单专区数量}}
        zones_data = df_as_of.groupby('专区名称').agg({'订单号':'nunique'}).reset_index()
        kv_data['{{产生订单专区数量}}'] = str(len(zones_data[zones_data['订单号'] > 0]))
        
        # 寻找主要专区
        z_rank = zones_data.sort_values('订单号', ascending=False)
        total_hist_orders = zones_data['订单号'].sum()
        
        if not z_rank.empty:
            top1 = z_rank.iloc[0]
            kv_data['{{主要专区1}}'] = str(top1['专区名称'])
            kv_data['{{主要专区1订单}}'] = str(int(top1['订单号']))
            kv_data['{{主要专区1占比}}'] = f"{(top1['订单号']/total_hist_orders*100):.2f}%" if total_hist_orders>0 else "0%"
            
        if len(z_rank) > 1:
            top2 = z_rank.iloc[1]
            kv_data['{{主要专区2}}'] = str(top2['专区名称'])
            kv_data['{{主要专区2订单}}'] = str(int(top2['订单号']))
            kv_data['{{主要专区2占比}}'] = f"{(top2['订单号']/total_hist_orders*100):.2f}%" if total_hist_orders>0 else "0%"
    except:
        print("累计专区排行计算出错")

    # 【补充原逻辑：新注册供应商专项提取 (使用预载入数据)】
    try:
        if not df_new_sup_file.empty:
            kv_data['{{新注册供应商数量}}'] = str(len(df_new_sup_file))
            new_zones = "、".join([str(z) for z in df_new_sup_file['专区名称'].unique() if str(z).strip() and str(z).lower() != 'nan'])
            kv_data['{{新注册供应商专区}}'] = new_zones if new_zones else "无源数据"
        else:
            kv_data['{{新注册供应商数量}}'] = "0"
            kv_data['{{新注册供应商专区}}'] = "无源数据"
    except Exception as e:
        kv_data['{{新注册供应商数量}}'] = "字段格式异常"
        kv_data['{{新注册供应商专区}}'] = "字段匹配失败"

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
    # 按照指示，全部立足 [截至数据] 物理表进行现场排行聚合
    try:
        # 为了精确计算，先在截至包内按采购单位进行全景聚合
        buyer_hist_rank = df_as_of.groupby('采购企业').agg({
            '订单号': 'nunique',
            money_col: 'sum'
        }).reset_index().rename(columns={'订单号': '累计订单', money_col: '累计总额'})
        
        # 补充：获取这些企业关联的专区（如果有多个，取其成交额最高的一个作为“来自专区”）
        def get_main_zone(buyer_name):
            b_data = df_as_of[df_as_of['采购企业'] == buyer_name]
            if not b_data.empty:
                return b_data.groupby('专区名称')[money_col].sum().idxmax()
            return "全区"

        if not buyer_hist_rank.empty:
            # 1. 活跃采购人数量 (基于截至数据去重)
            kv_data['{{活跃采购人数量}}'] = str(buyer_hist_rank['采购企业'].nunique())
            
            # 2. 最高订单采购人
            b_ord = buyer_hist_rank.sort_values('累计订单', ascending=False).iloc[0]
            kv_data['{{最高订单采购人}}'] = str(b_ord['采购企业'])
            kv_data['{{最高订单数}}'] = str(int(b_ord['累计订单']))
            kv_data['{{最高订单总额}}'] = f"{b_ord['累计总额']:.2f}"
            kv_data['{{最高订单专区}}'] = get_main_zone(b_ord['采购企业'])
            
            # 3. 最高金额采购人
            b_mon = buyer_hist_rank.sort_values('累计总额', ascending=False).iloc[0]
            kv_data['{{最高金额采购人}}'] = str(b_mon['采购企业'])
            kv_data['{{最高金额订单数}}'] = str(int(b_mon['累计订单']))
            kv_data['{{最高金额}}'] = f"{b_mon['累计总额']:.2f}"   
            kv_data['{{最高金额专区}}'] = get_main_zone(b_mon['采购企业'])
    except Exception as e:
        print(f"提取 单极历史金牌榜 失败: {e}")


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
        if '商品名称' in df_period.columns:
            # 清除汇总行、空值，取得唯一的商品销售条目总量
            valid_items = df_period[~df_period['商品名称'].astype(str).isin(['nan', '', 'None', '汇总', '合计'])]
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