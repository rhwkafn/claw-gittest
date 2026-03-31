# -*- coding: utf-8 -*-
"""
项目配置文件
"""

from dataclasses import dataclass
from typing import List, Dict
from datetime import datetime

@dataclass
class ReportConfig:
    """报告配置"""
    # 基本信息
    report_title: str = "2025 年华南地区新能源双碳目标工作报告"
    report_subtitle: str = "—— 广东、广西、海南三省区新能源发展综述"
    organization: str = "华南新能源发展研究中心"
    
    # 报告期间
    year: int = 2025
    start_month: int = 1
    end_month: int = 12
    
    # 覆盖区域
    regions: List[str] = None
    
    def __post_init__(self):
        if self.regions is None:
            self.regions = ['广东', '广西', '海南']
    
    @property
    def generated_date(self) -> str:
        return datetime.now().strftime('%Y 年 %-m 月 %-d 日')


@dataclass
class DataConfig:
    """数据配置"""
    # 数据单位
    unit: str = "亿千瓦时"
    
    # 月度数据（可以在这里替换为真实数据）
    guangdong_data: List[float] = None
    guangxi_data: List[float] = None
    hainan_data: List[float] = None
    
    def __post_init__(self):
        if self.guangdong_data is None:
            self.guangdong_data = [45.2, 42.8, 52.1, 58.6, 65.3, 72.1, 
                                   78.5, 76.2, 68.4, 58.9, 48.7, 44.5]
        if self.guangxi_data is None:
            self.guangxi_data = [28.3, 26.5, 35.2, 42.1, 48.6, 52.3,
                                 55.8, 53.2, 45.6, 38.4, 30.2, 27.8]
        if self.hainan_data is None:
            self.hainan_data = [12.5, 13.2, 18.6, 22.4, 26.8, 28.5,
                                30.2, 29.8, 25.6, 20.3, 15.8, 13.5]
    
    def validate(self) -> Dict[str, bool]:
        """验证数据完整性"""
        validation = {
            'guangdong_length': len(self.guangdong_data) == 12,
            'guangxi_length': len(self.guangxi_data) == 12,
            'hainan_length': len(self.hainan_data) == 12,
            'guangdong_positive': all(x > 0 for x in self.guangdong_data),
            'guangxi_positive': all(x > 0 for x in self.guangxi_data),
            'hainan_positive': all(x > 0 for x in self.hainan_data),
        }
        return validation
    
    def get_total(self) -> List[float]:
        """计算月度总计"""
        return [g + gx + h for g, gx, h in 
                zip(self.guangdong_data, self.guangxi_data, self.hainan_data)]
    
    def get_annual_total(self) -> Dict[str, float]:
        """计算年度总计"""
        return {
            '广东': sum(self.guangdong_data),
            '广西': sum(self.guangxi_data),
            '海南': sum(self.hainan_data),
            '总计': sum(self.get_total())
        }


@dataclass
class ChartConfig:
    """图表配置"""
    # 图片尺寸
    line_chart_size: tuple = (14, 7)
    pie_chart_size: tuple = (10, 8)
    bar_chart_size: tuple = (14, 7)
    
    # DPI
    dpi: int = 150
    
    # 颜色方案
    colors: Dict[str, str] = None
    
    def __post_init__(self):
        if self.colors is None:
            self.colors = {
                '广东': '#FF6B6B',
                '广西': '#4ECDC4',
                '海南': '#45B7D1',
                '华南总计': '#2C3E50'
            }


# 默认配置实例
DEFAULT_REPORT_CONFIG = ReportConfig()
DEFAULT_DATA_CONFIG = DataConfig()
DEFAULT_CHART_CONFIG = ChartConfig()