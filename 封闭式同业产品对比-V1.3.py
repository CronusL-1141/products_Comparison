#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# ============================================================================================
# Script Name   : 封闭式同业产品对比-V1.3.py
# Description   : 自动检测当前目录下的所有文件夹，每个文件夹中包含需对比的产品历史净值Excel文件；
#                 提取并对齐历史净值，进行插值补全，绘制“法巴农银”系列与其他产品对比图；
#                 支持根目录加载产品查询表，根据最新净值表中的“登记编码”筛选产品信息，
#                 并在输出Excel中新增“产品信息”表。
# Author        : Liu Guangjun
# Created Date  : 2025-06-13
# Last Modified : 2025-06-19
# Version       : v1.3
# Python Ver.   : 3.10+
# Dependencies  : pandas, re, colorsys, sys, os, xlsxwriter
# License       : Internal Use Only © 2025 法巴农银理财
# ============================================================================================

import pandas as pd
import os
import re
import sys
import colorsys

# 获取脚本根路径
if getattr(sys, 'frozen', False):
    root_dir = os.path.dirname(sys.executable)
else:
    root_dir = os.path.dirname(os.path.abspath(__file__))
print(f"📁 脚本根路径：{root_dir}")

# ———— 加载产品查询表 ————
files = os.listdir(root_dir)
query_file = next((f for f in files if f.endswith('.xlsx') and '产品查询' in f), None)
if query_file:
    query_path = os.path.join(root_dir, query_file)
    try:
        df_query = pd.read_excel(query_path, sheet_name='产品列表', header=8)
        print(f"📄 读取产品查询表：{query_file}")
    except Exception as e:
        print(f"❌ 读取产品查询表失败：{e}")
        df_query = None
else:
    print("⚠️ 未找到产品查询表，后续将不生成产品信息表。")
    df_query = None

def extract_product_name(filename):
    """从文件名中提取产品名称"""
    name = os.path.splitext(filename)[0]
    match = re.search(r'[\u4e00-\u9fa5A-Za-z0-9（）()]+$', name)
    return match.group(0) if match else name

# ———— 批量处理每个子文件夹 ————
folders = [d for d in os.listdir(root_dir) if os.path.isdir(os.path.join(root_dir, d))]
for folder in folders:
    folder_path = os.path.join(root_dir, folder)
    print(f"\n📂 处理文件夹：{folder}")
    file_list = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    product_nav = {}
    codes_in_folder = set()

    for file_name in file_list:
        file_path = os.path.join(folder_path, file_name)
        prod_name = extract_product_name(file_name)
        try:
            xls = pd.ExcelFile(file_path)
            # — 历史净值 —
            sheet_hist = next((s for s in xls.sheet_names if '历史净值' in s or '历史' in s), None)
            if sheet_hist:
                df_hist = xls.parse(sheet_hist)
                date_col = next((c for c in df_hist.columns if '日期' in c), None)
                nav_col  = next((c for c in df_hist.columns if '单位净值' in c), None)
                if date_col and nav_col:
                    tmp = df_hist[[date_col, nav_col]].dropna()
                    tmp.columns = ['date', prod_name]
                    tmp['date'] = pd.to_datetime(tmp['date'], errors='coerce')
                    tmp[prod_name] = pd.to_numeric(tmp[prod_name], errors='coerce')
                    tmp = tmp.dropna(subset=['date', prod_name]).groupby('date').last()
                    product_nav[prod_name] = tmp
                else:
                    print(f"⚠️ {file_name} 未找到“日期”或“单位净值”列，跳过历史数据。")
            else:
                print(f"⚠️ {file_name} 无历史净值Sheet，跳过。")

            # — 最新净值：提取“登记编码” —
            if df_query is not None:
                sheet_latest = next((s for s in xls.sheet_names if '最新净值' in s or '最新' in s), None)
                if sheet_latest:
                    df_latest = xls.parse(sheet_latest, header=0)
                    if '登记编码' in df_latest.columns:
                        codes = df_latest['登记编码'].dropna().astype(str).tolist()
                        codes_in_folder.update(codes)
                    else:
                        print(f"⚠️ {file_name} 最新净值表中无“登记编码”列。")
                else:
                    print(f"⚠️ {file_name} 无最新净值Sheet，无法提取登记编码。")

        except Exception as e:
            print(f"❌ 读取 {file_name} 出错：{e}")

    # — 合并、插值、作图准备 —
    if product_nav:
        merged_nav = pd.concat(product_nav.values(), axis=1).sort_index()
        merged_nav.index.name = '净值日期'
        plot_nav   = merged_nav.copy()
        full_idx   = pd.date_range(start=plot_nav.index.min(), end=plot_nav.index.max(), freq='D')
        plot_nav   = plot_nav.reindex(full_idx).interpolate(method='linear')
        plot_nav.index.name = '净值日期'

        # — 写入Excel —
        out_name = f"{folder}_产品净值汇总_含连接图表.xlsx"
        out_path = os.path.join(root_dir, out_name)
        with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
            # 历史净值
            df_out = merged_nav.reset_index()
            df_out['净值日期'] = df_out['净值日期'].dt.strftime('%Y年%m月%d日')
            df_out.to_excel(writer, sheet_name='净值数据', index=False)

            # 作图数据
            df_plot = plot_nav.reset_index()
            df_plot['净值日期'] = df_plot['净值日期'].dt.strftime('%Y-%m-%d')
            df_plot.to_excel(writer, sheet_name='作图数据', index=False)

            # 产品信息
            if df_query is not None and codes_in_folder:
                df_info = df_query[df_query['登记编码'].astype(str).isin(codes_in_folder)]
                df_info.to_excel(writer, sheet_name='产品信息', index=False)

            # 插入对比图
            workbook  = writer.book
            worksheet = writer.sheets['净值数据']
            chart     = workbook.add_chart({'type': 'line'})
            cols      = df_plot.columns[1:]
            fab_list  = [c for c in cols if '法巴农银' in c]
            other     = [c for c in cols if c not in fab_list]

            def gen_green(n):
                h = 120/360
                return [
                    '#{:02X}{:02X}{:02X}'.format(*[
                        int(c*255) for c in colorsys.hsv_to_rgb(
                            h, 1-0.4*(i/max(n-1,1)), 0.502+0.4*(i/max(n-1,1))
                        )
                    ]) for i in range(n)
                ]

            green_colors = gen_green(len(fab_list))
            non_green    = ['#5B9BD5', '#ED7D31', '#FFC000', '#4472C4',
                            '#A5A5A5', '#FF6666', '#8E44AD', '#2C3E50']

            for i, name in enumerate(fab_list):
                idx = list(cols).index(name) + 1
                chart.add_series({
                    'name':       ['作图数据', 0, idx],
                    'categories': ['作图数据', 1, 0, len(df_plot), 0],
                    'values':     ['作图数据', 1, idx, len(df_plot), idx],
                    'line':       {'color': green_colors[i], 'width': 2.0}
                })
            for i, name in enumerate(other):
                idx = list(cols).index(name) + 1
                chart.add_series({
                    'name':       ['作图数据', 0, idx],
                    'categories': ['作图数据', 1, 0, len(df_plot), 0],
                    'values':     ['作图数据', 1, idx, len(df_plot), idx],
                    'line':       {'color': non_green[i % len(non_green)], 'width': 2.0}
                })

            chart.set_title ({'name': '产品单位净值对比图'})
            chart.set_x_axis({'name': '日期', 'label_position': 'low'})
            chart.set_y_axis({'name': '单位净值'})
            chart.set_legend({'position': 'bottom'})
            chart.set_size  ({'width': 900, 'height': 400})
            worksheet.insert_chart('H2', chart)

        print(f"✅ 生成文件：{out_path}")
    else:
        print(f"⚠️ 文件夹 {folder} 未读取到净值数据，已跳过。")
