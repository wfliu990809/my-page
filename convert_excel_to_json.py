#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据转换脚本：读取所有Excel文件，提取日期和保费数据
"""

import pandas as pd
import json
import os
import re
import glob
from datetime import datetime


def extract_end_date_from_b1(filepath):
    """从B1单元格提取止期（结束日期）"""
    try:
        df = pd.read_excel(filepath, header=None)
        b1_content = str(df.iloc[0, 1])  # B1单元格
        
        # 匹配日期范围：统计日期:YYYY-MM-DD~YYYY-MM-DD
        date_match = re.search(r'统计日期:(\d{4}-\d{2}-\d{2})~(\d{4}-\d{2}-\d{2})', b1_content)
        if date_match:
            end_date = date_match.group(2)  # 提取止期
            return end_date
        return None
    except Exception as e:
        print(f"  - 提取日期失败: {e}")
        return None


def read_excel_data(filepath):
    """读取Excel文件的数据部分（从第3行开始）"""
    try:
        df = pd.read_excel(filepath, header=1)  # 第2行是表头
        df.columns = ['empty', '序号', '险种大类', '四级机构', '保费收入不含税']
        df = df[['序号', '险种大类', '四级机构', '保费收入不含税']]
        df = df.dropna(subset=['险种大类'])
        df = df[df['险种大类'] != '总计']
        df['保费收入不含税'] = pd.to_numeric(df['保费收入不含税'], errors='coerce').fillna(0)
        
        records = []
        for _, row in df.iterrows():
            records.append({
                'category': str(row['险种大类']).strip(),
                'institution': str(row['四级机构']).strip(),
                'premium': float(row['保费收入不含税'])
            })
        return records
    except Exception as e:
        print(f"  - 读取数据失败: {e}")
        return []


def determine_year_from_filename(filename):
    """根据文件名判断年份（2025或2026）"""
    if '2025' in filename:
        return 2025
    elif '2026' in filename:
        return 2026
    return None


def find_json_end(content, start_pos):
    """找到JSON对象的结束位置（考虑嵌套）"""
    brace_count = 0
    in_string = False
    escape_next = False
    
    for i in range(start_pos, len(content)):
        char = content[i]
        
        if escape_next:
            escape_next = False
            continue
        
        if char == '\\':
            escape_next = True
            continue
        
        if char == '"' and not in_string:
            in_string = True
        elif char == '"' and in_string:
            in_string = False
        elif not in_string:
            if char == '{':
                brace_count += 1
            elif char == '}':
                brace_count -= 1
                if brace_count == 0:
                    return i
    
    return -1


def update_html_embedded_data(json_data):
    """更新HTML文件中的内嵌数据"""
    html_file = 'index.html'
    
    if not os.path.exists(html_file):
        print(f"警告: {html_file} 不存在，跳过更新内嵌数据")
        return False
    
    with open(html_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    json_str = json.dumps(json_data, ensure_ascii=False, separators=(',', ':'))
    
    # 找到EMBEDDED_DATA的位置
    start_marker = 'const EMBEDDED_DATA = '
    start_pos = content.find(start_marker)
    
    if start_pos == -1:
        print(f"警告: 未找到 EMBEDDED_DATA 标记")
        return False
    
    data_start = start_pos + len(start_marker)
    json_end = find_json_end(content, data_start)
    
    if json_end == -1:
        print(f"警告: 无法找到JSON结束位置")
        return False
    
    # 构建新HTML：替换JSON数据（保留后面的分号）
    new_content = content[:data_start] + json_str + content[json_end+1:]
    
    with open(html_file, 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    print(f"  - 已更新 {html_file} 中的内嵌数据")
    return True


def main():
    """主函数"""
    
    # 数据文件夹路径
    data_dir = r'D:\程序库\业务平台分析\源数据'
    
    # 如果文件夹不存在，提示创建
    if not os.path.exists(data_dir):
        print(f'错误: 数据文件夹不存在: {data_dir}')
        print('请创建该文件夹并将Excel文件放入其中')
        return
    
    # 扫描所有Excel文件
    excel_files = []
    for pattern in ['*.xlsx', '*.xls']:
        excel_files.extend(glob.glob(os.path.join(data_dir, pattern)))
    
    # 过滤掉生成的文件
    excel_files = [f for f in excel_files if not f.startswith('~')]
    
    print(f"发现 {len(excel_files)} 个Excel文件")
    
    # 数据结构：按年份和日期存储
    data_by_year = {
        2025: {},  # date -> {records: [], total: 0}
        2026: {}
    }
    
    all_categories = set()
    all_institutions = set()
    
    for filepath in excel_files:
        filename = os.path.basename(filepath)
        print(f"\n处理: {filename}")
        
        # 判断年份
        year = determine_year_from_filename(filename)
        if year is None:
            print(f"  - 跳过：无法从文件名识别年份")
            continue
        
        # 提取日期
        end_date = extract_end_date_from_b1(filepath)
        if end_date is None:
            print(f"  - 跳过：无法提取日期")
            continue
        
        print(f"  - 年份: {year}, 日期: {end_date}")
        
        # 读取数据
        records = read_excel_data(filepath)
        if not records:
            print(f"  - 跳过：无数据")
            continue
        
        # 计算总保费
        total_premium = sum(r['premium'] for r in records)
        
        # 收集类别和机构
        for r in records:
            all_categories.add(r['category'])
            all_institutions.add(r['institution'])
        
        # 存储数据（如果同一日期已有数据，提示覆盖，不要累加）
        if end_date in data_by_year[year]:
            print(f"  - 警告: {end_date} 已有数据，将被覆盖为: {filename}")
        
        data_by_year[year][end_date] = {
            'date': end_date,
            'filename': filename,
            'records': records,
            'total_premium': total_premium
        }
        
        print(f"  - 成功读取 {len(records)} 条记录，累计保费: {total_premium:,.2f}")
    
    # 准备图表数据：按日期排序的时间序列
    chart_data = {
        2025: [],
        2026: []
    }
    
    for year in [2025, 2026]:
        dates = sorted(data_by_year[year].keys())
        for date in dates:
            chart_data[year].append({
                'date': date,
                'total_premium': data_by_year[year][date]['total_premium'],
                'filename': data_by_year[year][date]['filename']
            })
    
    # 构建最终结果
    result = {
        'metadata': {
            'generated_at': datetime.now().isoformat(),
            'categories': sorted(list(all_categories)),
            'institutions': sorted(list(all_institutions)),
            'dates_2025': sorted(data_by_year[2025].keys()),
            'dates_2026': sorted(data_by_year[2026].keys())
        },
        'time_series': chart_data,
        'detail_data': {
            '2025': data_by_year[2025],
            '2026': data_by_year[2026]
        }
    }
    
    # 保存为JSON文件
    with open('data.json', 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    # 更新HTML中的内嵌数据
    print("\n正在更新HTML文件...")
    update_html_embedded_data(result)
    
    print(f"\n[OK] 数据转换完成！")
    print(f"  - JSON文件: data.json")
    print(f"  - 2025年数据点: {len(chart_data[2025])} 个")
    print(f"  - 2026年数据点: {len(chart_data[2026])} 个")
    print(f"  - 险种类别: {len(all_categories)} 个")
    print(f"  - 四级机构: {len(all_institutions)} 个")
    print(f"\n现在可以直接双击打开 index.html 查看数据看板")


if __name__ == '__main__':
    main()
