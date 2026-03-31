# -*- coding: utf-8 -*-
"""
2015-2025 年华南地区新能源数据趋势分析
生成可视化图表和统计报告
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')

plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# 读取数据
data_path = r"D:\AI-agent\openclaw-apri\projects\2025-新能源华南双碳报告\crawler\data\2015-2025 年华南地区新能源历史数据.csv"
df = pd.read_csv(data_path)

print("\n=== 加载数据完成 ===")
print(f"数据行数：{len(df)}")
print(f"年份范围：{df['年份'].min()} - {df['年份'].max()}")

# 打印列名
print("\n列名:", df.columns.tolist())

# 使用正确的列名
numeric_cols = [col for col in df.columns if col.endswith('亿千瓦时')]
print(f"数值列：{numeric_cols}")

# 年度汇总
yearly = df.groupby('年份')[numeric_cols].sum().round(1)
print("\n【年度汇总】（单位：亿千瓦时）")
print(yearly)

# ========== 图表生成 ==========
output_dir = r"D:\AI-agent\openclaw-apri\projects\2025-新能源华南双碳报告\charts"

# 使用 iloc 获取列（避免编码问题）
total_col = numeric_cols[-1]  # 新能源合计是最后一列

# 图 1: 年度总发电量趋势
fig1, ax1 = plt.subplots(figsize=(12, 6))
years = yearly.index.astype(str)
ax1.plot(years, yearly[total_col], marker='o', linewidth=2, label='新能源合计', color='#2C3E50')

for i, (year, val) in enumerate(zip(years, yearly[total_col])):
    ax1.annotate(f'{val:.1f}', (i, val), textcoords="offset points", xytext=(0,5), ha='center')

ax1.set_xlabel('年份', fontsize=12)
ax1.set_ylabel('发电量 (亿千瓦时)', fontsize=12)
ax1.set_title('2015-2025 年华南地区新能源年度发电量趋势', fontsize=14)
ax1.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig(f"{output_dir}/年度趋势_2015-2025.png", dpi=150)
plt.close()
print(f"\n[OK] 已保存：年度趋势_2015-2025.png")

# 图 2: 分能源类型年度对比
fig2, ax2 = plt.subplots(figsize=(14, 7))
x = years
width = 0.2

energy_types = numeric_cols[:4]  # 前 4 个是风电/太阳能/水电/核电
labels = ['风电', '太阳能', '水电', '核电']
colors = ['#FF6B6B', '#FFD93D', '#4ECDC4', '#6B5B95']

for i, (etype, label, color) in enumerate(zip(energy_types, labels, colors)):
    values = [yearly[etype].iloc[j] for j in range(len(yearly))]
    ax2.bar([float(xi) + i*width for xi in x], values, width, label=label, color=color)

ax2.set_xlabel('年份', fontsize=12)
ax2.set_ylabel('发电量 (亿千瓦时)', fontsize=12)
ax2.set_title('2015-2025 年华南地区新能源分类型年度发电量', fontsize=14)
ax2.set_xticks([float(xi) + 1.5*width for xi in x])
ax2.set_xticklabels(years)
ax2.legend()
ax2.grid(True, alpha=0.3, axis='y')
plt.tight_layout()
plt.savefig(f"{output_dir}/分能源类型对比_2015-2025.png", dpi=150)
plt.close()
print(f"[OK] 已保存：分能源类型对比_2015-2025.png")

# 图 3: 分省年度对比
fig3, ax3 = plt.subplots(figsize=(14, 7))
provinces = df['省份'].unique()
colors_prov = ['#FF6B6B', '#4ECDC4', '#45B7D1']

for i, (prov, color) in enumerate(zip(provinces, colors_prov)):
    prov_data = df[df['省份'] == prov].groupby('年份')[total_col].sum()
    values = [prov_data.get(year, 0) for year in range(2015, 2026)]
    ax3.plot(years, values, marker='o', linewidth=2, label=prov, color=color)

ax3.set_xlabel('年份', fontsize=12)
ax3.set_ylabel('发电量 (亿千瓦时)', fontsize=12)
ax3.set_title('2015-2025 年华南三省新能源年度发电量对比', fontsize=14)
ax3.legend()
ax3.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig(f"{output_dir}/分省对比_2015-2025.png", dpi=150)
plt.close()
print(f"[OK] 已保存：分省对比_2015-2025.png")

# ========== 统计摘要 ==========
summary = {
    'total_2015_2025': float(yearly[total_col].sum()),
    'growth_rate': float((yearly[total_col].iloc[-1] - yearly[total_col].iloc[0]) / yearly[total_col].iloc[0] * 100),
    'avg_annual_growth': float(yearly[total_col].pct_change().mean() * 100),
}

print("\n=== 趋势分析摘要 ===")
print(f"2015-2025 年累计发电量：{summary['total_2015_2025']:.1f} 亿千瓦时")
print(f"总增长率：{summary['growth_rate']:.1f}%")
print(f"年均增长率：{summary['avg_annual_growth']:.2f}%")

print("\n[OK] 趋势分析完成！")