# 2025 年华南地区新能源双碳目标工作报告生成系统

📊 自动化生成新能源数据分析报告和可视化图表

---

## 📁 项目结构

```
2025-新能源华南双碳报告/
├── config.py              # 配置文件（v2.0 新增）
├── generate_data.py       # 数据生成与可视化脚本
├── generate_report.py     # Word 报告生成脚本
├── README.md              # 项目说明
├── 工作流程总结.md         # 工作流程文档
├── .gitignore             # Git 忽略配置
│
├── data/                  # 数据文件
│   └── 2025 年华南新能源发电量.xlsx
│
├── charts/                # 可视化图表
│   ├── 月度发电量折线图.png
│   ├── 全年发电量占比饼图.png
│   └── 各省月度对比柱状图.png
│
└── reports/               # 报告文档
    └── 2025 年华南地区新能源双碳目标工作报告.docx
```

---

## 🚀 快速开始

### 环境要求

- Python 3.8+
- 中文字体（微软雅黑、黑体、楷体）

### 安装依赖

```bash
pip install pandas openpyxl matplotlib python-docx
```

### 运行脚本

```bash
# 生成数据和图表
python generate_data.py

# 生成 Word 报告
python generate_report.py
```

---

## ⚙️ 配置说明 (v2.0+)

编辑 `config.py` 自定义：

### 报告配置
```python
ReportConfig(
    report_title="你的报告标题",
    organization="你的单位名称",
    year=2025
)
```

### 数据配置
```python
DataConfig(
    guangdong_data=[...],  # 12 个月数据
    guangxi_data=[...],
    hainan_data=[...]
)
```

### 图表配置
```python
ChartConfig(
    dpi=150,  # 图片分辨率
    colors={'广东': '#FF6B6B', ...}  # 颜色方案
)
```

---

## 📊 功能特性

### v1.0
- ✅ 月度发电量数据生成
- ✅ Excel 数据表格（带样式）
- ✅ 3 种可视化图表（折线图/饼图/柱状图）
- ✅ Word 报告自动生成

### v2.0 (当前版本)
- ✅ 配置化设计（config.py）
- ✅ 数据验证功能
- ✅ 可扩展的数据源
- ✅ 自定义颜色方案
- ✅ 模块化代码结构

---

## 📈 输出示例

### 数据文件
- Excel 表格包含月度数据、年度汇总
- 自动计算各省占比

### 可视化图表
| 图表类型 | 说明 |
|---------|------|
| 折线图 | 月度趋势，显示季节性变化 |
| 饼图 | 三省区年度贡献占比 |
| 柱状图 | 月度对比，直观显示差异 |

### Word 报告
- 专业封面和目录
- 7 个完整章节
- 中文排版优化

---

## 🔧 自定义数据

可以替换为真实数据源：

```python
# 方法 1: 直接修改 config.py
DataConfig(
    guangdong_data=[真实数据...],
)

# 方法 2: 从 Excel/CSV 导入
import pandas as pd
df = pd.read_excel('真实数据.xlsx')
```

---

## 📝 版本历史

| 版本 | 日期 | 更新内容 |
|------|------|---------|
| v2.0 | 2026-04-01 | 添加配置化、数据验证 |
| v1.0 | 2026-04-01 | 初始版本 |

---

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

---

## 📄 许可证

MIT License

---

**开发时间:** 2026 年 4 月  
**工作区:** `D:\AI-agent\openclaw-apri\projects\`