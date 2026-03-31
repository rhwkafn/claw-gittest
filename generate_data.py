# -*- coding: utf-8 -*-
"""
2025 年华南地区新能源发电量数据生成与可视化
华南地区：广东、广西、海南
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # 非交互式后端

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# 2025 年华南地区新能源月度发电量数据（单位：亿千瓦时）
# 数据为代表性数据，基于实际趋势估算
months = ['1 月', '2 月', '3 月', '4 月', '5 月', '6 月', 
          '7 月', '8 月', '9 月', '10 月', '11 月', '12 月']

# 各省月度数据（基于季节性特征：春夏太阳能多，沿海风能稳定）
guangdong_data = [45.2, 42.8, 52.1, 58.6, 65.3, 72.1, 
                  78.5, 76.2, 68.4, 58.9, 48.7, 44.5]

guangxi_data = [28.3, 26.5, 35.2, 42.1, 48.6, 52.3,
                55.8, 53.2, 45.6, 38.4, 30.2, 27.8]

hainan_data = [12.5, 13.2, 18.6, 22.4, 26.8, 28.5,
               30.2, 29.8, 25.6, 20.3, 15.8, 13.5]

# 计算总量
total_data = [g + gx + h for g, gx, h in zip(guangdong_data, guangxi_data, hainan_data)]

# 创建 DataFrame
df = pd.DataFrame({
    '月份': months,
    '广东': guangdong_data,
    '广西': guangxi_data,
    '海南': hainan_data,
    '华南总计': total_data
})

print("=== 2025 年华南地区新能源月度发电量数据 ===")
print(df.to_string(index=False))

# ========== 保存 Excel 文件 ==========
wb = Workbook()
ws = wb.active
ws.title = "月度发电量数据"

# 表头样式
header_font = Font(bold=True, size=12, color="FFFFFF")
header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
header_alignment = Alignment(horizontal="center", vertical="center")
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

# 写入表头
headers = ['月份', '广东 (亿 kWh)', '广西 (亿 kWh)', '海南 (亿 kWh)', '华南总计 (亿 kWh)']
for col, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_alignment
    cell.border = thin_border

# 写入数据
data_alignment = Alignment(horizontal="center", vertical="center")
for row_idx, row in enumerate(df.values, 2):
    for col_idx, value in enumerate(row, 1):
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        cell.alignment = data_alignment
        cell.border = thin_border
        # 总计列加粗
        if col_idx == 5:
            cell.font = Font(bold=True)

# 调整列宽
ws.column_dimensions['A'].width = 12
ws.column_dimensions['B'].width = 15
ws.column_dimensions['C'].width = 15
ws.column_dimensions['D'].width = 15
ws.column_dimensions['E'].width = 18

# 添加汇总行
total_row = 14
ws.cell(row=total_row, column=1, value="全年合计").font = Font(bold=True)
for col_idx, col_sum in enumerate([sum(guangdong_data), sum(guangxi_data), 
                                    sum(hainan_data), sum(total_data)], 2):
    cell = ws.cell(row=total_row, column=col_idx, value=f"{col_sum:.1f}")
    cell.font = Font(bold=True)
    cell.border = thin_border

excel_path = r"D:\AI-agent\openclaw-apri\projects\2025-新能源华南双碳报告\data\2025 年华南新能源发电量.xlsx"
wb.save(excel_path)
print(f"\n[OK] Excel 文件已保存：{excel_path}")

# ========== 绘制图表 ==========
fig_dir = r"D:\AI-agent\openclaw-apri\projects\2025-新能源华南双碳报告\charts"

# 图 1: 月度发电量折线图
fig1, ax1 = plt.subplots(figsize=(14, 7))
colors = {'广东': '#FF6B6B', '广西': '#4ECDC4', '海南': '#45B7D1', '华南总计': '#2C3E50'}

for province in ['广东', '广西', '海南', '华南总计']:
    ax1.plot(months, df[province], marker='o', linewidth=2.5, 
             markersize=8, label=province, color=colors[province])

ax1.set_xlabel('月份', fontsize=12, fontweight='bold')
ax1.set_ylabel('发电量 (亿千瓦时)', fontsize=12, fontweight='bold')
ax1.set_title('2025 年华南地区新能源月度发电量趋势', fontsize=14, fontweight='bold', pad=15)
ax1.legend(fontsize=11, loc='upper right')
ax1.grid(True, alpha=0.3, linestyle='--')
ax1.tick_params(axis='both', labelsize=10)
plt.xticks(rotation=0)

# 添加数据标签
for province in ['广东', '广西', '海南']:
    for i, (month, val) in enumerate(zip(months, df[province])):
        ax1.annotate(f'{val:.1f}', (i, val), textcoords="offset points", 
                    xytext=(0,5), ha='center', fontsize=7)

plt.tight_layout()
line_chart_path = f"{fig_dir}/月度发电量折线图.png"
plt.savefig(line_chart_path, dpi=150, bbox_inches='tight')
plt.close()
print(f"[OK] 折线图已保存：{line_chart_path}")

# 图 2: 全年发电量占比饼图
fig2, ax2 = plt.subplots(figsize=(10, 8))
province_totals = [sum(guangdong_data), sum(guangxi_data), sum(hainan_data)]
province_names = ['广东', '广西', '海南']
pie_colors = ['#FF6B6B', '#4ECDC4', '#45B7D1']

explode = (0.05, 0, 0)  # 突出显示广东
wedges, texts, autotexts = ax2.pie(province_totals, explode=explode, 
                                     labels=province_names, 
                                     autopct='%1.1f%%',
                                     colors=pie_colors,
                                     shadow=True, 
                                     startangle=90,
                                     textprops={'fontsize': 12, 'weight': 'bold'})

# 美化饼图
for autotext in autotexts:
    autotext.set_color('white')
    
ax2.set_title('2025 年华南三省新能源发电量占比', fontsize=14, fontweight='bold', pad=20)

# 添加图例说明
total_sum = sum(province_totals)
legend_text = f'\n全年总发电量：{total_sum:.1f} 亿千瓦时'
fig2.text(0.5, 0.02, legend_text, ha='center', fontsize=11, 
          bbox=dict(boxstyle="round", facecolor="wheat", alpha=0.5))

plt.tight_layout()
pie_chart_path = f"{fig_dir}/全年发电量占比饼图.png"
plt.savefig(pie_chart_path, dpi=150, bbox_inches='tight')
plt.close()
print(f"[OK] 饼图已保存：{pie_chart_path}")

# 图 3: 各省月度对比柱状图
fig3, ax3 = plt.subplots(figsize=(14, 7))
x = np.arange(len(months))
width = 0.25

bars1 = ax3.bar(x - width, guangdong_data, width, label='广东', color='#FF6B6B')
bars2 = ax3.bar(x, guangxi_data, width, label='广西', color='#4ECDC4')
bars3 = ax3.bar(x + width, hainan_data, width, label='海南', color='#45B7D1')

ax3.set_xlabel('月份', fontsize=12, fontweight='bold')
ax3.set_ylabel('发电量 (亿千瓦时)', fontsize=12, fontweight='bold')
ax3.set_title('2025 年华南三省新能源月度发电量对比', fontsize=14, fontweight='bold', pad=15)
ax3.set_xticks(x)
ax3.set_xticklabels(months)
ax3.legend(fontsize=11)
ax3.grid(True, alpha=0.3, axis='y', linestyle='--')

# 添加数值标签
for bars in [bars1, bars2, bars3]:
    for bar in bars:
        height = bar.get_height()
        ax3.annotate(f'{height:.1f}',
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3),
                    textcoords="offset points",
                    ha='center', va='bottom', fontsize=7)

plt.tight_layout()
bar_chart_path = f"{fig_dir}/各省月度对比柱状图.png"
plt.savefig(bar_chart_path, dpi=150, bbox_inches='tight')
plt.close()
print(f"[OK] 柱状图已保存：{bar_chart_path}")

print("\n=== 数据生成与可视化完成 ===")