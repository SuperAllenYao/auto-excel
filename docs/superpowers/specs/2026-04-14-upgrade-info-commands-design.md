# 设计文档：upgrade 和 info 命令

## 概述

为 auto-excel CLI 新增两个命令：

- **`info`**：显示版本号、安装路径、Python 版本、数据目录等运行环境信息
- **`upgrade`**：自动从 GitHub 拉取最新版本并安装升级，支持版本对比和"已是最新"检测

## 背景

- 现有命令：`version`（仅版本号）、`on`（处理文件）、`uninstall`（卸载）
- 安装机制：`install.sh` 将代码克隆到 `~/.auto-excel`，通过 wrapper 脚本 `~/.local/bin/auto-excel` 调用
- 升级逻辑 `install.sh` 已有：`git pull` + `uv sync`
- 版本存储：`src/auto_excel/__init__.py` 中 `__version__ = "1.1.1"`

## 路径常量抽取

`uninstall`（`cli.py:79-80`）硬编码了安装路径。新增的 `upgrade` 和 `info` 也需要这些路径，共 3 处复用。将其提升到 `config.py`：

```python
# config.py 新增
INSTALL_DIR = Path.home() / ".auto-excel"
WRAPPER = Path.home() / ".local" / "bin" / "auto-excel"
```

`uninstall` 命令改为引用 `config.INSTALL_DIR` 和 `config.WRAPPER`。

## info 命令

### 输出格式

```
auto-excel 1.1.1
安装路径: /Users/allen/.auto-excel
Python:   3.11.9
数据目录: /Users/allen/Desktop/marketing analysis
```

### 实现

- 版本：`from auto_excel import __version__`
- 安装路径：`config.INSTALL_DIR`
- Python 版本：`sys.version_info`
- 数据目录：`config.BASE_DIR`
- 输出方式：`typer.echo`，与现有 `version` 命令风格一致

### 与 version 命令的关系

`version` 命令保持不变，`info` 是更详细的环境信息展示。

## upgrade 命令

### 完整流程

```
用户执行 auto-excel upgrade
  │
  ├── 1. 检查 INSTALL_DIR/.git 是否存在
  │      └── 不存在 → 报错并显示重新安装命令 → 退出
  │
  ├── 2. 读取并显示当前版本: "当前版本: v1.1.1"
  │
  ├── 3. git -C INSTALL_DIR fetch origin master
  │      └── 失败 → 报错: "无法连接到远程仓库，请检查网络" → 退出
  │
  ├── 4. 对比 git rev-parse HEAD vs git rev-parse origin/master
  │      └── 相同 → "已是最新版本 v1.1.1，无需升级" → 退出
  │
  ├── 5. git -C INSTALL_DIR pull origin master
  │      └── 失败 → 显示 git 错误输出 → 退出
  │
  ├── 6. cd INSTALL_DIR && uv sync
  │      └── 失败 → 显示 uv 错误输出 → 退出
  │
  ├── 7. 从文件系统解析新版本号
  │      读取 INSTALL_DIR/src/auto_excel/__init__.py
  │      正则提取 __version__ = "x.y.z"
  │
  └── 8. 显示: "升级成功: v1.1.1 → v1.2.0"
```

### 新版本读取

升级后 Python 模块缓存仍为旧版本。通过直接读取文件系统并正则解析获取新版本：

```python
import re
text = (config.INSTALL_DIR / "src" / "auto_excel" / "__init__.py").read_text()
match = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', text)
new_version = match.group(1) if match else "unknown"
```

### subprocess 调用

所有外部命令通过 `subprocess.run()` 执行，`capture_output=True`，`text=True`。检查 `returncode` 判断成败。

## 错误处理

| 错误场景 | 处理方式 |
|----------|----------|
| `~/.auto-excel` 不存在或非 git 仓库 | 提示重新安装，显示 `curl ... \| bash` 命令 |
| `git` 不在 PATH | 提示 "请先安装 git" |
| `uv` 不在 PATH | 提示 "请先安装 uv" |
| 网络失败（fetch/pull） | 显示 git 错误输出 |
| `uv sync` 失败 | 显示 uv 错误输出 |

## 测试策略

### info 命令

- `CliRunner` 调用 `["info"]`
- 验证输出包含版本号（`__version__`）
- 验证输出包含安装路径字符串
- 验证输出包含 Python 版本字符串

### upgrade 命令

通过 `monkeypatch` mock `subprocess.run`，覆盖三条路径：

1. **已是最新版**：fetch 成功，两个 rev-parse 返回相同 hash → 输出包含 "已是最新"
2. **成功升级**：fetch 成功，hash 不同，pull 成功，sync 成功 → 输出包含 "升级成功" 和新旧版本号
3. **安装目录不存在**：INSTALL_DIR 不存在 → 输出包含错误提示

## 文件变更清单

| 文件 | 变更类型 | 说明 |
|------|----------|------|
| `src/auto_excel/config.py` | 修改 | 新增 `INSTALL_DIR`、`WRAPPER` 常量 |
| `src/auto_excel/cli.py` | 修改 | 新增 `info`、`upgrade` 命令；`uninstall` 改用 config 常量 |
| `tests/test_cli.py` | 修改 | 新增 info/upgrade 测试用例 |
