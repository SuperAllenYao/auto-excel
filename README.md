# auto-excel

macOS 专用的营销分析 Excel 自动化处理 CLI 工具。将原始 Excel 文件放入桌面指定文件夹，一条命令即可完成数据计算、排序、分组统计，并输出处理后的新文件。

## 功能特性

- **列计算**：自动插入「实际花费」「点击率」「CPC」「实际成本」四个计算列
- **智能排序**：按实际成本从高到低自动排序，表头固定
- **分组统计**：按高（≥90）/ 中（50–89）/ 低（<50）三档分组，合并单元格并显示占比
- **增量处理**：已处理的文件自动跳过，失败的文件下次自动重试
- **一键卸载**：`auto-excel uninstall` 彻底移除程序，保留已处理文档
- **友好输出**：Rich 终端实时进度与执行报告

## 安装

```bash
curl -sSL https://raw.githubusercontent.com/SuperAllenYao/auto-excel/master/install.sh | bash
```

安装完成后新开终端窗口，将 Excel 文件放入 `~/Desktop/marketing analysis/Raw/`，然后运行：

```bash
auto-excel on
```

## 目录结构

```
~/Desktop/marketing analysis/
├── Raw/          # 原始 Excel 文件（程序只读）
├── New/          # 处理后的输出文件（与原文件同名）
└── log/
    ├── processed.json    # 已处理文件状态记录
    └── YYYY-MM-DD.log    # 运行日志
```

## 处理流程

程序自动对第 4 个 Sheet 执行三步处理：

1. **列计算**（在指定列右侧插入）
   - 实际花费 = 花费 ÷ 1.136
   - 点击率 = 点击量 ÷ 展现量
   - CPC = 实际花费 ÷ 点击量
   - 实际成本 = 实际花费 ÷ 留资人数

2. **排序**：按实际成本从高到低排序

3. **分组统计**：在「互动成本」右侧插入「占比」列，按高 / 中 / 低三档分组合并单元格

## 技术栈

| 层级 | 选型 |
|------|------|
| 语言 | Python 3.11+ |
| 包管理 | uv |
| Excel 处理 | openpyxl 3.1.5+ |
| CLI 框架 | Typer 0.9+ |
| 终端输出 | Rich 14.0+ |

## 版本

v1.1.0 — 新增 uninstall 命令
