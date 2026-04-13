# 更新日志

## [v1.1.0] - 2026-04-13

### ✨ 新增

- **uninstall 命令**：`auto-excel uninstall` 一键卸载，删除 `~/.auto-excel/` 和 `~/.local/bin/auto-excel`，执行前显示确认提示，**不删除**已处理文档（`~/Desktop/marketing analysis/` 完整保留）

### 🐛 修复

- **版本号显示**：`auto-excel version` 现在从 `__version__` 读取，不再显示硬编码的旧版本号
- **安装脚本 URL**：修正 `README.md` 中 `curl` 安装命令的分支名（`main` → `master`）

---

## [v1.0.0] - 2026-04-13

### ✨ 新增

- **CLI 入口**：`auto-excel on` 命令，自动处理 `Raw/` 目录下所有未处理的 `.xlsx` 文件
- **Flow 1 列计算**：在第 4 个 Sheet 中插入「实际花费」「点击率」「CPC」「实际成本」四个计算列；分母为零时填 0 并记录日志
- **Flow 2 排序**：按「实际成本」列从高到低排序数据行，表头固定不动
- **Flow 3 分组统计**：按高（≥90）/ 中（50–89）/ 低（<50）三档分组，合并「占比」列单元格并填入「条数/占比」
- **状态管理**：`log/processed.json` 记录已处理文件，重复运行自动跳过；处理失败的文件不写入状态，下次自动重试
- **Rich 终端输出**：实时进度提示、执行报告表格，全中文 + emoji 展示
- **一键安装脚本** `install.sh`：自动安装 uv、同步依赖、创建 wrapper、配置 PATH、建立桌面工作目录；幂等，可重复运行
- **74 个自动化测试**：涵盖单元测试、对抗性测试（Mutation Survivor、Boundary Assassin 等六类）及端对端集成测试
