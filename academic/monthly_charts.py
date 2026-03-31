# -*- coding: utf-8 -*-
"""
月度数据可视化 - 学术级图表生成
2015-2025 年华南地区新能源月度分析
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')
from matplotlib.gridspec import GridSpec
from scipy import stats
import os

plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['figure.dpi'] = 150

# 路径
BASE = r"D:\AI-agent\openclaw-apri\projects\2025-新能源华南双碳报告"
DATA_PATH = f"{BASE}\\crawler\\data\\2015-2025 年华南地区新能源历史数据.csv"
OUT_DIR = f"{BASE}\\academic\\figures"
os.makedirs(OUT_DIR, exist_ok=True)

df = pd.read_csv(DATA_PATH)
numeric_cols = [col for col in df.columns if col.endswith('亿千瓦时')]
total_col = numeric_cols[-1]
energy_cols = numeric_cols[:4]
energy_labels = ['风电', '太阳能', '水电', '核电']
energy_colors = ['#1F77B4', '#FF7F0E', '#2CA02C', '#D62728']
prov_colors = {'广东': '#E74C3C', '广西': '#2ECC71', '海南': '#3498DB'}

df_total = df.groupby(['年份', '月份'])[numeric_cols].sum().reset_index()
df_total['时间'] = df_total['年份'].astype(str) + '-' + df_total['月份'].astype(str).str.zfill(2)
df_total = df_total.sort_values(['年份', '月份'])

# =============================================
# 图 1: 全序列月度总发电量时间轴
# =============================================
fig, ax = plt.subplots(figsize=(18, 6))
time_index = range(len(df_total))
ax.fill_between(time_index, df_total[total_col], alpha=0.3, color='#2C3E50')
ax.plot(time_index, df_total[total_col], linewidth=1.5, color='#2C3E50')

for year in range(2015, 2026):
    idx = df_total[df_total['年份'] == year].index
    if len(idx) > 0:
        pos = list(df_total['年份']).index(year)
        ax.axvline(x=pos, color='gray', linestyle='--', alpha=0.4, linewidth=0.8)
        ax.text(pos + 0.5, df_total[total_col].max() * 0.95, str(year), fontsize=8, color='gray')

ax.set_xlabel('月份序列（2015.01 - 2025.12）', fontsize=12)
ax.set_ylabel('发电量（亿千瓦时）', fontsize=12)
ax.set_title('2015-2025 年华南地区新能源月度发电量时序图（Fig.1）', fontsize=13, fontweight='bold')
ax.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig(f"{OUT_DIR}/Fig1_月度时序图.png", dpi=150, bbox_inches='tight')
plt.close()
print("[OK] Fig1 月度时序图")

# =============================================
# 图 2: 各能源类型月度发电量堆叠面积图
# =============================================
fig, ax = plt.subplots(figsize=(18, 7))
bottom = np.zeros(len(df_total))
for i, (col, label, color) in enumerate(zip(energy_cols, energy_labels, energy_colors)):
    ax.fill_between(time_index, bottom, bottom + df_total[col].values,
                    label=label, color=color, alpha=0.8)
    bottom += df_total[col].values

for year in range(2015, 2026):
    pos = list(df_total['年份']).index(year) if year in df_total['年份'].values else None
    if pos is not None:
        ax.axvline(x=pos, color='white', linestyle='--', alpha=0.5, linewidth=0.8)

ax.set_xlabel('月份序列（2015.01 - 2025.12）', fontsize=12)
ax.set_ylabel('发电量（亿千瓦时）', fontsize=12)
ax.set_title('2015-2025 年华南地区各类新能源月度发电量结构（Fig.2）', fontsize=13, fontweight='bold')
ax.legend(loc='upper left', fontsize=11)
ax.grid(True, alpha=0.2)
plt.tight_layout()
plt.savefig(f"{OUT_DIR}/Fig2_能源结构堆叠图.png", dpi=150, bbox_inches='tight')
plt.close()
print("[OK] Fig2 能源结构堆叠图")

# =============================================
# 图 3: 月度季节性热力图
# =============================================
fig, axes = plt.subplots(1, 2, figsize=(16, 7))
years = sorted(df_total['年份'].unique())
months = list(range(1, 13))
month_labels = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

heatmap_data = pd.pivot_table(df_total, values=total_col, index='月份', columns='年份')
im1 = axes[0].imshow(heatmap_data.values, cmap='YlOrRd', aspect='auto')
axes[0].set_xticks(range(len(years)))
axes[0].set_xticklabels(years, rotation=45)
axes[0].set_yticks(range(12))
axes[0].set_yticklabels(month_labels)
axes[0].set_title('月度发电量热力图（Fig.3a）', fontsize=12, fontweight='bold')
plt.colorbar(im1, ax=axes[0], label='亿千瓦时')

monthly_avg = df_total.groupby('月份')[total_col].mean()
seasonal_index = monthly_avg / monthly_avg.mean() * 100
axes[1].bar(month_labels, seasonal_index.values, color=energy_colors[2], alpha=0.8, edgecolor='white')
axes[1].axhline(y=100, color='red', linestyle='--', linewidth=1.5, label='基准线（SI=100）')
for i, v in enumerate(seasonal_index.values):
    axes[1].text(i, v + 0.5, f'{v:.1f}', ha='center', fontsize=9)
axes[1].set_xlabel('月份', fontsize=11)
axes[1].set_ylabel('季节性指数 (SI)', fontsize=11)
axes[1].set_title('月度季节性指数（Fig.3b）', fontsize=12, fontweight='bold')
axes[1].legend()
axes[1].grid(True, alpha=0.3, axis='y')
plt.tight_layout()
plt.savefig(f"{OUT_DIR}/Fig3_季节性分析.png", dpi=150, bbox_inches='tight')
plt.close()
print("[OK] Fig3 季节性分析")

# =============================================
# 图 4: 分省月度对比（近 3 年箱线图）
# =============================================
fig, axes = plt.subplots(1, 3, figsize=(18, 6))
recent_df = df[df['年份'] >= 2023]

for ax, (prov, color) in zip(axes, prov_colors.items()):
    prov_data = recent_df[recent_df['省份'] == prov]
    monthly_data = [prov_data[prov_data['月份'] == m][total_col].values for m in range(1, 13)]
    bp = ax.boxplot(monthly_data, patch_artist=True,
                   medianprops=dict(color='white', linewidth=2))
    for patch in bp['boxes']:
        patch.set_facecolor(color)
        patch.set_alpha(0.7)
    ax.set_xticklabels(['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'], rotation=45, fontsize=8)
    ax.set_title(f'{prov}（2023-2025）（Fig.4）', fontsize=11, fontweight='bold')
    ax.set_ylabel('月度发电量（亿千瓦时）')
    ax.grid(True, alpha=0.3, axis='y')
plt.suptitle('华南三省新能源月度发电量分布（2023-2025）', fontsize=13, fontweight='bold')
plt.tight_layout()
plt.savefig(f"{OUT_DIR}/Fig4_分省月度箱线图.png", dpi=150, bbox_inches='tight')
plt.close()
print("[OK] Fig4 分省月度箱线图")

# =============================================
# 图 5: 年均增速与线性趋势
# =============================================
fig, ax = plt.subplots(figsize=(12, 6))
yearly = df_total.groupby('年份')[total_col].sum()
years_arr = np.array(yearly.index)
values_arr = np.array(yearly.values)
slope, intercept, r, p, se = stats.linregress(years_arr, values_arr)
trend_line = slope * years_arr + intercept

ax.bar(years_arr, values_arr, color='#3498DB', alpha=0.7, label='实际发电量')
ax.plot(years_arr, trend_line, 'r--', linewidth=2,
        label=f'线性趋势 (y={slope:.1f}x+{intercept:.0f}, R²={r**2:.3f})')

for i, (year, val) in enumerate(zip(years_arr, values_arr)):
    if i > 0:
        growth = (val - values_arr[i-1]) / values_arr[i-1] * 100
        ax.text(year, val + 100, f'+{growth:.1f}%', ha='center', fontsize=8, color='green')

ax.set_xlabel('年份', fontsize=12)
ax.set_ylabel('年度发电量（亿千瓦时）', fontsize=12)
ax.set_title('2015-2025 年华南地区新能源年度发电量及增长趋势（Fig.5）', fontsize=13, fontweight='bold')
ax.legend(fontsize=11)
ax.grid(True, alpha=0.3, axis='y')
plt.tight_layout()
plt.savefig(f"{OUT_DIR}/Fig5_年度增长趋势.png", dpi=150, bbox_inches='tight')
plt.close()
print("[OK] Fig5 年度增长趋势")

# =============================================
# 图 6: 各能源类型占比演变
# =============================================
fig, ax = plt.subplots(figsize=(14, 7))
yearly_by_type = df.groupby('年份')[energy_cols].sum()
for col, label, color in zip(energy_cols, energy_labels, energy_colors):
    share = yearly_by_type[col] / yearly_by_type[energy_cols].sum(axis=1) * 100
    ax.plot(yearly_by_type.index, share, marker='o', linewidth=2, label=label, color=color)
    ax.fill_between(yearly_by_type.index, share, alpha=0.1, color=color)

ax.set_xlabel('年份', fontsize=12)
ax.set_ylabel('占比（%）', fontsize=12)
ax.set_title('2015-2025 年华南地区各类新能源发电量占比演变（Fig.6）', fontsize=13, fontweight='bold')
ax.legend(fontsize=11)
ax.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig(f"{OUT_DIR}/Fig6_能源占比演变.png", dpi=150, bbox_inches='tight')
plt.close()
print("[OK] Fig6 能源占比演变")

print("\n[OK] 全部学术图表已生成！")
print(f"图表路径：{OUT_DIR}")