# -*- coding: utf-8 -*-
"""
数据分析模块
提供统计分析、趋势预测、数据导出等功能

v3.0 新增
"""

import json
import pandas as pd
import numpy as np
from datetime import datetime
from typing import Dict, List, Any
from config import DataConfig, DEFAULT_DATA_CONFIG


class DataAnalyzer:
    """数据分析器"""
    
    def __init__(self, data_config: DataConfig = None):
        self.data_config = data_config or DEFAULT_DATA_CONFIG
        self.months = ['1 月', '2 月', '3 月', '4 月', '5 月', '6 月', 
                       '7 月', '8 月', '9 月', '10 月', '11 月', '12 月']
    
    def get_statistics(self) -> Dict[str, Any]:
        """计算统计指标"""
        total = self.data_config.get_total()
        annual = self.data_config.get_annual_total()
        
        stats = {
            '广东': {
                '年总计': annual['广东'],
                '月平均': np.mean(self.data_config.guangdong_data),
                '最大值': max(self.data_config.guangdong_data),
                '最小值': min(self.data_config.guangdong_data),
                '标准差': np.std(self.data_config.guangdong_data),
                '峰值月份': self.months[np.argmax(self.data_config.guangdong_data)],
                '谷值月份': self.months[np.argmin(self.data_config.guangdong_data)],
            },
            '广西': {
                '年总计': annual['广西'],
                '月平均': np.mean(self.data_config.guangxi_data),
                '最大值': max(self.data_config.guangxi_data),
                '最小值': min(self.data_config.guangxi_data),
                '标准差': np.std(self.data_config.guangxi_data),
                '峰值月份': self.months[np.argmax(self.data_config.guangxi_data)],
                '谷值月份': self.months[np.argmin(self.data_config.guangxi_data)],
            },
            '海南': {
                '年总计': annual['海南'],
                '月平均': np.mean(self.data_config.hainan_data),
                '最大值': max(self.data_config.hainan_data),
                '最小值': min(self.data_config.hainan_data),
                '标准差': np.std(self.data_config.hainan_data),
                '峰值月份': self.months[np.argmax(self.data_config.hainan_data)],
                '谷值月份': self.months[np.argmin(self.data_config.hainan_data)],
            },
            '总计': {
                '年总计': annual['总计'],
                '月平均': np.mean(total),
                '最大值': max(total),
                '最小值': min(total),
                '峰值月份': self.months[np.argmax(total)],
                '谷值月份': self.months[np.argmin(total)],
            }
        }
        
        # 添加占比
        for province in ['广东', '广西', '海南']:
            stats[province]['占比'] = round(annual[province] / annual['总计'] * 100, 2)
        
        return stats
    
    def get_growth_analysis(self) -> Dict[str, Any]:
        """增长分析（环比）"""
        data = {
            '广东': self.data_config.guangdong_data,
            '广西': self.data_config.guangxi_data,
            '海南': self.data_config.hainan_data,
        }
        
        growth = {}
        for province, values in data.items():
            monthly_growth = []
            for i in range(1, len(values)):
                if values[i-1] > 0:
                    growth_rate = (values[i] - values[i-1]) / values[i-1] * 100
                    monthly_growth.append(round(growth_rate, 2))
                else:
                    monthly_growth.append(None)
            
            growth[province] = {
                '月度环比': ['N/A'] + monthly_growth,
                '平均增长率': round(np.mean([g for g in monthly_growth if g is not None]), 2) if monthly_growth else 0,
                '最高增长月份': self.months[monthly_growth.index(max([g for g in monthly_growth if g is not None])) + 1] if monthly_growth else 'N/A',
            }
        
        return growth
    
    def get_seasonal_analysis(self) -> Dict[str, Any]:
        """季节性分析"""
        seasons = {
            '春季': [2, 3, 4],   # 3-5 月
            '夏季': [5, 6, 7],   # 6-8 月
            '秋季': [8, 9, 10],  # 9-11 月
            '冬季': [11, 0, 1],  # 12-2 月
        }
        
        seasonal = {}
        data = {
            '广东': self.data_config.guangdong_data,
            '广西': self.data_config.guangxi_data,
            '海南': self.data_config.hainan_data,
        }
        
        for province, values in data.items():
            seasonal[province] = {}
            for season_name, months_idx in seasons.items():
                season_total = sum(values[i] for i in months_idx)
                season_avg = season_total / 3
                seasonal[province][season_name] = round(season_avg, 2)
        
        return seasonal
    
    def get_carbon_reduction(self) -> Dict[str, Any]:
        """碳减排估算"""
        # 按照中国电网平均排放因子：约 0.8 吨 CO2/MWh = 0.08 万吨/亿 kWh
        emission_factor = 0.8  # 吨 CO2/MWh
        
        annual = self.data_config.get_annual_total()
        
        carbon = {}
        for province in ['广东', '广西', '海南', '总计']:
            # 亿 kWh = 1000 万 MWh
            mwh = annual[province] * 10  # 转换为万 MWh
            co2_reduction = mwh * emission_factor / 10000  # 万吨
            carbon[province] = {
                '发电量 (亿 kWh)': annual[province],
                '碳减排量 (万吨 CO2)': round(co2_reduction, 2),
                '等效植树 (万棵)': round(co2_reduction * 55, 2),  # 约 55 棵树/吨 CO2/年
            }
        
        return carbon
    
    def export_json(self, output_path: str) -> None:
        """导出完整分析结果为 JSON"""
        result = {
            'meta': {
                'report_title': '2025 年华南地区新能源数据分析报告',
                'generated_at': datetime.now().isoformat(),
                'data_source': 'config.py',
                'version': '3.0',
            },
            'statistics': self.get_statistics(),
            'growth_analysis': self.get_growth_analysis(),
            'seasonal_analysis': self.get_seasonal_analysis(),
            'carbon_reduction': self.get_carbon_reduction(),
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"[OK] 分析结果已导出：{output_path}")
    
    def print_summary(self) -> None:
        """打印分析报告摘要"""
        print("\n" + "=" * 70)
        print("2025 年华南地区新能源数据分析报告 v3.0")
        print("=" * 70)
        
        stats = self.get_statistics()
        carbon = self.get_carbon_reduction()
        
        print("\n【发电量统计】")
        print("-" * 70)
        for province in ['广东', '广西', '海南']:
            s = stats[province]
            print(f"\n{province}:")
            print(f"  年总计：{s['年总计']:.1f} 亿千瓦时 (占比 {s['占比']}%)")
            print(f"  月平均：{s['月平均']:.1f} 亿千瓦时")
            print(f"  峰值：{s['最大值']:.1f} 亿千瓦时 ({s['峰值月份']})")
            print(f"  谷值：{s['最小值']:.1f} 亿千瓦时 ({s['谷值月份']})")
        
        print(f"\n【华南总计】{stats['总计']['年总计']:.1f} 亿千瓦时")
        print(f"   月平均：{stats['总计']['月平均']:.1f} 亿千瓦时")
        print(f"   峰值月份：{stats['总计']['峰值月份']}")
        print(f"   谷值月份：{stats['总计']['谷值月份']}")
        
        print("\n【碳减排效益】")
        print("-" * 70)
        for province in ['广东', '广西', '海南', '总计']:
            c = carbon[province]
            print(f"  {province}: 减排 {c['碳减排量 (万吨 CO2)']:.2f} 万吨 CO2 "
                  f"(等效植树 {c['等效植树 (万棵)']:.2f} 万棵)")
        
        print("\n" + "=" * 70)


def main():
    """主函数"""
    analyzer = DataAnalyzer()
    analyzer.print_summary()
    
    # 导出 JSON
    output_path = r"D:\AI-agent\openclaw-apri\projects\2025-新能源华南双碳报告\data\analysis_result.json"
    analyzer.export_json(output_path)
    
    print("\n[OK] 分析完成")


if __name__ == '__main__':
    main()