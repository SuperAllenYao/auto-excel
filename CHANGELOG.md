# 更新日志

## [v2.0.0] - 2026-04-16

### ✨ 新增

- **公式感知预处理**：新增 `resolve_formulas(wb, ws)` 函数，按三阶段解析 Excel 公式串
  - SUMIFS 公式：扫描目标表公式 → 读取源表 → 按 key 聚合 → 回填数值
  - 除法公式：`=X/Y` 在 SUMIFS 解析后计算，零除与非数值统一兜底为 0
  - 残余清理：任何剩余的 `=…` 字符串最终替换为 0，防止下游 `apply_calculated_columns` 接触公式串
- **空行过滤**：新增 `remove_empty_rows(ws)` 函数，倒序删除「笔记标题 + 笔记ID」均为空的行，保护 header，避免空行污染分组统计
- **流水线接入**：`process_file` 新增两个中文进度回调 `正在解析公式...` / `正在过滤空行...`，在计算列之前运行

### 🔄 变更

- **Excel 读取模式**：`load_workbook(..., data_only=True)` 改为 `data_only=False`，保留公式字符串供 `resolve_formulas` 解析；其他工作表的公式也得以完整保存到输出文件
- **SUMIFS 正则捕获**：源 sheet 的 key 列字母改为从公式中动态捕获（原先硬编码 A），对公式层级更鲁棒

### 🐛 修复

- **公式列读不到数据的历史问题**：此前 `data_only=True` 在 openpyxl 下导致公式缓存缺失时读取为 0 或 None，`apply_calculated_columns` 的下游计算全部失真；现已由 `resolve_formulas` 显式求值
- **整表公式被抹除**：`data_only=False` 模式下其他 sheet 的公式不再被覆写为空
- **openpyxl hyperlink.ref 不随 delete_rows 同步**：`remove_empty_rows` 删行后，单元格 hyperlink 的 `ref` 仍指向旧坐标，导致保存后文件重载出现 25 行幽灵空行；修复为删除循环结束后统一重置 `hyperlink.ref = cell.coordinate`
- **空键但有公式的行被误清零**：Phase 4 新增 `row_has_formula` 判断，完全空行保留 None 交由 `remove_empty_rows` 处理，仅对「有公式但 key 为空」的行填 0
- **除零/非数值兜底**：除法阶段统一捕获 `ZeroDivisionError / TypeError / ValueError` → 0

### ⚠️ 破坏性变更

- **底层读取模式从 `data_only=True` 切到 `data_only=False`**：对于没有公式的纯数值工作簿，行为与之前一致（向后兼容测试 `test_full_pipeline` 保留）；对于含公式的工作簿，本版改由内置解析器负责，不再依赖 Excel 缓存的计算值。若外部脚本直接复用 `process_file` 且依赖 openpyxl 读取到已缓存的公式结果，升级前请核对

### ✅ 测试

- 测试数量：109 → **157**（新增 48 条，含单元、集成、对抗测试，3 阶段审查覆盖每项任务）
- 新增对抗用例：None/空串 key 碰撞、每行除法独立、Phase 5/6 在无 SUMIFS 时仍必须运行、hyperlink 回归保护等
- 真实文件端到端验证：406 行输入 → 100 行纯净输出，0 幽灵空行

---

## [v1.2.0] - 2026-04-14

### ✨ 新增

- **info 命令**：`auto-excel info` 显示版本号、安装路径、Python 版本、数据目录等环境信息
- **upgrade 命令**：`auto-excel upgrade` 自动从 GitHub 检测并升级到最新版本，支持"已是最新"检测和升级前后版本对比

### 🔄 变更

- **安装路径常量化**：将 `INSTALL_DIR` 和 `WRAPPER` 提取到 `config.py`，`uninstall`/`upgrade`/`info` 三个命令共享

---

## [v1.1.1] - 2026-04-13

### 🐛 修复

- **公式单元格兼容**：源文件中含 Excel 公式（如 `=SUMIFS(...)`）的列不再导致处理失败；改用 `data_only=True` 读取缓存计算值，并对无法转换为数字的值记录警告并填 0

---

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
