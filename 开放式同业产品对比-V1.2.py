#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# =============================================================================
# Script Name   : 开放式同业产品对比-V1.2.py
# Description   : 批量读取各子文件夹或当前目录下的产品历史净值Excel文件，
#                 提取并统一对齐至法巴产品起始日，
#                 对其他产品净值进行归一化处理；
#                 支持根目录加载产品查询表，根据最新净值表中的“登记编码”
#                 筛选产品信息并输出至“产品信息”Sheet（可自定义列顺序）。
# Author        : Liu Guangjun
# Created Date  : 2025-06-20
# Version       : v2.0
# Python Ver.   : 3.10+
# Dependencies  : pandas, re, colorsys, sys, os, xlsxwriter
# License       : Internal Use Only © 2025 法巴农银理财
# =============================================================================

import pandas as pd
import os
import re
import sys
import colorsys
from datetime import datetime

# —————— 配置 ——————
# 自定义产品信息表列顺序
desired_info_cols = [
    '理财产品名称','最早实际成立日期','最早实际结束日期','投资周期（天）',
    '业绩比较基准（%）','最新销售费(%)','最新固定管理费(%)','折合人民币计算日期存续规模',
    '近1个月年化收益率(%)','近3个月年化收益率(%)','成立以来年化收益率(%)',
    '近1个月最大回撤(%)','近3个月最大回撤(%)','成立以来最大回撤(%)'
]

# 获取根目录（.exe或.py）
if getattr(sys, 'frozen', False):
    root_dir = os.path.dirname(sys.executable)
else:
    root_dir = os.path.dirname(os.path.abspath(__file__))
print(f"📁 根目录：{root_dir}")

# —————— 加载产品查询表 ——————
files = os.listdir(root_dir)
query_file = next((f for f in files if f.endswith('.xlsx') and '产品查询' in f), None)
if query_file:
    try:
        df_query = pd.read_excel(
            os.path.join(root_dir, query_file),
            sheet_name='产品列表', header=8
        )
        print(f"📄 加载产品查询表：{query_file}")
    except Exception as e:
        print(f"❌ 读取产品查询表失败：{e}")
        df_query = None
else:
    print("⚠️ 未找到产品查询表，不生成产品信息表。")
    df_query = None

# — 提取文件夹列表（若无子文件夹则处理当前目录）
subdirs = [d for d in os.listdir(root_dir) if os.path.isdir(os.path.join(root_dir, d))]
work_dirs = subdirs if subdirs else [root_dir]

# 遍历每个工作目录
for work in work_dirs:
    work_path = os.path.join(root_dir, work)
    print(f"\n🚀 处理目录：{work_path}")

    # 查找所有.xlsx文件（排除临时文件）
    excel_files = [f for f in os.listdir(work_path)
                   if f.endswith('.xlsx') and not f.startswith('~$')]
    product_nav = {}
    codes_in_folder = set()

    # 遍历Excel文件
    for fname in excel_files:
        path = os.path.join(work_path, fname)
        # 提取产品名称
        name = os.path.splitext(fname)[0]
        m = re.search(r'[一-龥A-Za-z0-9（）()]+$', name)
        prod = m.group(0).strip() if m else name.strip()

        try:
            xls = pd.ExcelFile(path)
            # 历史净值
            sheet_h = next((s for s in xls.sheet_names if '历史净值' in s or '历史' in s), None)
            if sheet_h:
                df_h = xls.parse(sheet_h)
                dc = next((c for c in df_h.columns if '日期' in c), None)
                nc = next((c for c in df_h.columns if '单位净值' in c), None)
                if dc and nc:
                    t = df_h[[dc,nc]].dropna()
                    t.columns = ['date', prod]
                    t['date'] = pd.to_datetime(t['date'], errors='coerce')
                    t[prod] = pd.to_numeric(t[prod], errors='coerce')
                    t = t.dropna(subset=['date', prod]).groupby('date').last()
                    product_nav[prod] = t
            # 最新净值 -> 提取登记编码
            if df_query is not None:
                sheet_l = next((s for s in xls.sheet_names if '最新净值' in s or '最新' in s), None)
                if sheet_l:
                    df_l = xls.parse(sheet_l, header=0)
                    if '登记编码' in df_l.columns:
                        codes_in_folder.update(df_l['登记编码'].dropna().astype(str).tolist())
        except Exception as e:
            print(f"❌ 读取 {fname} 异常：{e}")

    # 若无数据则跳过
    if not product_nav:
        print(f"⚠️ 目录 {work} 无净值数据，跳过。")
        continue

    # 合并原始净值
    df_raw = pd.concat(product_nav.values(), axis=1).sort_index()
    df_raw.index.name = '净值日期'

    # 法巴系列起始日
    fab_cols = [c for c in df_raw.columns if '法巴' in c]
    start = df_raw[fab_cols].dropna(how='all').index.min()

    # 归一化非法巴产品
    df_norm = df_raw.copy()
    for col in df_raw.columns:
        if col in fab_cols: continue
        if start in df_raw.index and pd.notna(df_raw.at[start, col]):
            df_norm[col + '（统一基准）'] = df_raw[col] / df_raw.at[start, col]

    # 重排序
    order = []
    for c in df_raw.columns:
        order.append(c)
        nu = c + '（统一基准）'
        if nu in df_norm.columns:
            order.append(nu)
    df_norm = df_norm[order]
    df_norm.index.name = '净值日期'

    # 日度插值
    full_idx = pd.date_range(df_raw.index.min(), df_raw.index.max(), freq='D')
    df_plot_raw  = df_raw.reindex(full_idx).interpolate('linear')
    df_plot_norm = df_norm.reindex(full_idx).interpolate('linear')
    df_plot_norm = df_plot_norm[df_plot_norm.index >= start]

    # 准备输出
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    out_name = f"{work}_开放式产品对比_{ts}.xlsx"
    out_path = os.path.join(work_path, out_name)

    with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
        df_raw.reset_index().to_excel(writer, sheet_name='原始净值', index=False)
        df_norm.reset_index().to_excel(writer, sheet_name='统一基准净值', index=False)
        df_plot_raw.reset_index().to_excel(writer, sheet_name='原始净值作图', index=False)
        df_plot_norm.reset_index().to_excel(writer, sheet_name='统一基准作图', index=False)

        # 产品信息 Sheet
        if df_query is not None and codes_in_folder:
            df_info = df_query[df_query['登记编码'].astype(str).isin(codes_in_folder)].copy()
            cols1 = [c for c in desired_info_cols if c in df_info.columns]
            cols2 = [c for c in df_info.columns if c not in cols1]
            df_info = df_info[cols1 + cols2]
            df_info.to_excel(writer, sheet_name='产品信息', index=False)

        # 图表生成函数
        def gen_green(n):
            h = 120/360
            return ['#{:02X}{:02X}{:02X}'.format(*[int(c*255) for c in colorsys.hsv_to_rgb(
                h, 1-0.4*(i/max(n-1,1)), 0.502+0.4*(i/max(n-1,1))
            )]) for i in range(n)]

        other_cols = ['#5B9BD5','#ED7D31','#FFC000','#4472C4','#A5A5A5','#FF6666','#8E44AD','#2C3E50']
        def add_chart(sheet, df, title):
            chart = writer.book.add_chart({'type':'line'})
            names = df.columns[1:]
            fab = [n for n in names if n in fab_cols]
            oth = [n for n in names if n not in fab]
            green = gen_green(len(fab))
            for i,nm in enumerate(fab):
                idx = df.columns.get_loc(nm)
                chart.add_series({
                    'name':[sheet,0,idx],'categories':[sheet,1,0,len(df),0],
                    'values':[sheet,1,idx,len(df),idx],'line':{'color':green[i],'width':2.0}
                })
            for i,nm in enumerate(oth):
                idx = df.columns.get_loc(nm)
                chart.add_series({
                    'name':[sheet,0,idx],'categories':[sheet,1,0,len(df),0],
                    'values':[sheet,1,idx,len(df),idx],'line':{'color':other_cols[i%len(other_cols)],'width':2.0}
                })
            chart.set_title({'name':title})
            chart.set_x_axis({'name':'日期'})
            chart.set_y_axis({'name':'单位净值'})
            chart.set_legend({'position':'bottom'})
            chart.set_size({'width':900,'height':400})
            writer.sheets[sheet].insert_chart('H2', chart)

        add_chart('原始净值作图', df_plot_raw.reset_index(), '原始产品单位净值')
        add_chart('统一基准作图', df_plot_norm.reset_index(), '统一基准单位净值')

    print(f"✅ 输出：{out_path}")
