#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
导出独立版HTML - 将所有数据打包到一个文件中
别人无需Python环境，直接双击打开即可查看
"""

import json
import os
from datetime import datetime


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


def export_standalone_html():
    """导出独立版HTML文件"""
    
    # 读取当前HTML模板
    with open('index.html', 'r', encoding='utf-8') as f:
        html_content = f.read()
    
    # 读取最新数据
    if not os.path.exists('data.json'):
        print("错误: 未找到 data.json，请先运行 convert_excel_to_json.py")
        return False
    
    with open('data.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # 将数据转换为JSON字符串
    data_json = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
    
    # 找到EMBEDDED_DATA的位置
    start_marker = 'const EMBEDDED_DATA = '
    start_pos = html_content.find(start_marker)
    
    if start_pos == -1:
        print("错误: 未找到 EMBEDDED_DATA 标记")
        return False
    
    data_start = start_pos + len(start_marker)
    json_end = find_json_end(html_content, data_start)
    
    print(f"调试: 数据开始位置={data_start}, JSON结束位置={json_end}")
    
    if json_end == -1:
        print("错误: 无法找到JSON结束位置")
        return False
    
    # 获取旧的JSON内容用于比较
    old_json = html_content[data_start:json_end+1]
    print(f"调试: 旧JSON长度={len(old_json)}, 新JSON长度={len(data_json)}")
    print(f"调试: 内容是否相同={old_json == data_json}")
    
    # 构建新HTML
    new_html = html_content[:data_start] + data_json + html_content[json_end+1:]
    
    if new_html == html_content:
        print("警告: 内容未发生变化（JSON数据可能已经是最新的）")
        # 即使内容相同，也继续导出
    
    # 提取最新日期（从2026年数据中找最晚的日期）
    all_dates = []
    if data['time_series'].get('2026'):
        all_dates.extend([item['date'] for item in data['time_series']['2026']])
    if data['time_series'].get('2025'):
        all_dates.extend([item['date'] for item in data['time_series']['2025']])
    
    if all_dates:
        all_dates.sort()
        latest_date = all_dates[-1]  # 格式: 2026-03-04
        date_str = latest_date[5:].replace('-', '')  # 格式: 0304
    else:
        date_str = datetime.now().strftime('%m%d')
    
    # 生成带日期的文件名
    output_filename = f'保费数据监控看板_最新版（{date_str}）.html'
    
    # 添加导出标记注释
    export_info = f"""<!-- 
===========================================
保费数据监控看板 - 独立导出版
导出时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
数据来源: D:\\程序库\\业务平台分析\\源数据
包含文件: {len(data['time_series']['2025']) + len(data['time_series']['2026'])} 个时间点
===========================================
使用说明:
1. 此文件包含所有数据，无需其他文件
2. 直接双击打开即可查看
3. 支持筛选险种和机构
===========================================
-->\n"""
    
    new_html = export_info + new_html
    
    # 保存文件
    with open(output_filename, 'w', encoding='utf-8') as f:
        f.write(new_html)
    
    print(f"\n[OK] 导出成功！")
    print(f"   文件: {output_filename}")
    print(f"\n[数据摘要]")
    print(f"   - 2025年数据点: {len(data['time_series']['2025'])} 个")
    print(f"   - 2026年数据点: {len(data['time_series']['2026'])} 个")
    print(f"   - 险种类别: {len(data['metadata']['categories'])} 个")
    print(f"   - 四级机构: {len(data['metadata']['institutions'])} 个")
    print(f"\n[提示]")
    print(f"   将文件发送给别人，对方直接双击打开即可查看")
    print(f"   无需安装Python，无需任何其他文件")
    
    return True


def main():
    """主函数"""
    print("=" * 60)
    print("导出独立版HTML")
    print("=" * 60)
    print()
    
    # 检查是否需要先更新数据
    data_dir = r'D:\程序库\业务平台分析\源数据'
    if os.path.exists(data_dir):
        excel_files = [f for f in os.listdir(data_dir) if f.endswith(('.xlsx', '.xls'))]
        print(f"发现 {len(excel_files)} 个Excel文件在源数据文件夹")
        
        # 检查data.json是否存在且最新
        if os.path.exists('data.json'):
            data_mtime = os.path.getmtime('data.json')
            latest_excel_mtime = max(
                [os.path.getmtime(os.path.join(data_dir, f)) for f in excel_files],
                default=0
            )
            
            if latest_excel_mtime > data_mtime:
                print("\n[!] 检测到Excel文件有更新，建议先运行数据转换")
                print("   是否现在运行转换? (y/n): ", end='')
                choice = input().strip().lower()
                if choice == 'y':
                    print("\n正在运行数据转换...")
                    os.system('python convert_excel_to_json.py')
                    print()
    
    # 导出独立版
    if export_standalone_html():
        print("\n" + "=" * 60)
        print("导出完成！")
        print("=" * 60)
    else:
        print("\n[ERROR] 导出失败")


if __name__ == '__main__':
    main()
