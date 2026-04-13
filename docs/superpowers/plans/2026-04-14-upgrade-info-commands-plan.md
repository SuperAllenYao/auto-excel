# upgrade & info 命令实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers-pro:subagent-driven-development (recommended) or superpowers-pro:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 为 auto-excel CLI 新增 `upgrade`（自动升级）和 `info`（显示环境信息）命令，并将安装路径抽取为共享常量。

**Architecture:** 在 `config.py` 新增 `INSTALL_DIR`/`WRAPPER` 常量供 `uninstall`/`upgrade`/`info` 三个命令复用。`info` 直接读取本地信息输出。`upgrade` 通过 `subprocess.run` 调用 `git fetch/pull` + `uv sync` 完成升级，升级前通过 `rev-parse` 对比 commit hash 判断是否需要升级。

**Tech Stack:** Python 3.11+, Typer 0.9+, subprocess (标准库), pytest + monkeypatch

---

## Part 1: Requirement Specification

### Sub-function overview

1. **路径常量抽取** — `INSTALL_DIR` 和 `WRAPPER` 从 `uninstall` 提升到 `config.py`
2. **uninstall 重构** — 改用共享常量，行为不变
3. **info 命令** — 显示版本号、安装路径、Python 版本、数据目录
4. **upgrade 命令** — 从 GitHub 检测并拉取最新版，安装依赖，显示版本对比

### Sub-function details

#### info 命令

- **输出字段:**
  - 版本号：字符串，来自 `auto_excel.__version__`
  - 安装路径：Path，`config.INSTALL_DIR` 的字符串表示
  - Python 版本：`{major}.{minor}.{micro}` 格式
  - 数据目录：Path，`config.BASE_DIR` 的字符串表示
- **完整流程:** 用户执行 `auto-excel info` → 四行信息逐行打印 → 退出码 0
- **边界:** 所有字段为本地读取，无失败分支

#### upgrade 命令

- **完整流程:**
  1. 检查 `INSTALL_DIR/.git` 是否为目录 → 不是则报错（含重装命令），退出码 1
  2. 显示 "当前版本: v{current}"
  3. `git -C {INSTALL_DIR} fetch origin master` → returncode ≠ 0 则报错，退出码 1
  4. 对比 `git -C {INSTALL_DIR} rev-parse HEAD` vs `git -C {INSTALL_DIR} rev-parse origin/master` 的 stdout（去 whitespace）
  5. 相同 → 输出 "已是最新版本 v{current}，无需升级"，退出码 0
  6. 不同 → `git -C {INSTALL_DIR} pull origin master` → returncode ≠ 0 则报错，退出码 1
  7. `subprocess.run(["uv", "sync"], cwd=INSTALL_DIR)` → returncode ≠ 0 则报错，退出码 1
  8. 从 `INSTALL_DIR/src/auto_excel/__init__.py` 读取新版本
  9. 输出 "升级成功: v{old} → v{new}"，退出码 0
- **边界:**
  - 安装目录不存在：提示包含 `curl -sSL ... | bash` 重装命令
  - 网络失败：显示 git stderr
  - `uv sync` 失败：显示 uv stderr
  - 版本解析失败：新版本显示为 "unknown"

### Non-functional requirements

- 无性能要求（均为低频操作）
- 无安全要求（无认证、无敏感数据处理）

---

## File Structure

| File | Operation | Responsibility |
|------|-----------|----------------|
| `src/auto_excel/config.py` | Modify | 新增 `INSTALL_DIR`、`WRAPPER` 常量 |
| `src/auto_excel/cli.py` | Modify | 新增 `info()`、`upgrade()` 命令，`_parse_version_from_file()` 辅助函数；`uninstall()` 改用 config 常量 |
| `tests/test_cli.py` | Modify | 新增 config 常量、info、upgrade 全部测试用例 |

---

### Task 1: Add install path constants to config.py

**Type:** Code task

**Context Brief:**
- **Files:** `src/auto_excel/config.py` (modify), `tests/test_cli.py` (modify)
- **Interface:** Two module-level constants: `INSTALL_DIR: Path`, `WRAPPER: Path`
- **Constraints:** Follow existing naming style in config.py (UPPER_SNAKE_CASE Path constants). `INSTALL_DIR` = `Path.home() / ".auto-excel"`, `WRAPPER` = `Path.home() / ".local" / "bin" / "auto-excel"`. These values must match `install.sh` lines 15-17 and current `cli.py:79-80`.

**Files:**
- Modify: `src/auto_excel/config.py:13` (append after `STATE_FILE`)
- Test: `tests/test_cli.py` (append new tests at end)

- [ ] **Step 1: Write failing tests**

Add to the end of `tests/test_cli.py`:

```python
def test_config_install_dir_is_path():
    from pathlib import Path
    from auto_excel.config import INSTALL_DIR
    assert isinstance(INSTALL_DIR, Path)
    assert INSTALL_DIR.name == ".auto-excel"

def test_config_wrapper_is_path():
    from pathlib import Path
    from auto_excel.config import WRAPPER
    assert isinstance(WRAPPER, Path)
    assert WRAPPER.name == "auto-excel"
    assert ".local" in str(WRAPPER)
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest tests/test_cli.py::test_config_install_dir_is_path tests/test_cli.py::test_config_wrapper_is_path -v`
Expected: FAIL with `ImportError` — `INSTALL_DIR` and `WRAPPER` not found

- [ ] **Step 3: Implement constants**

In `src/auto_excel/config.py`, append after line 12 (`STATE_FILE = ...`):

```python
INSTALL_DIR = Path.home() / ".auto-excel"
WRAPPER = Path.home() / ".local" / "bin" / "auto-excel"
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest tests/test_cli.py::test_config_install_dir_is_path tests/test_cli.py::test_config_wrapper_is_path -v`
Expected: 2 PASSED

- [ ] **Step 5: Write adversarial tests**

Add to `tests/test_cli.py`:

```python
def test_adversarial_install_dir_under_home():
    from auto_excel.config import INSTALL_DIR
    from pathlib import Path
    assert str(INSTALL_DIR).startswith(str(Path.home()))

def test_adversarial_wrapper_under_local_bin():
    from auto_excel.config import WRAPPER
    assert WRAPPER.parent.name == "bin"
    assert WRAPPER.parent.parent.name == ".local"
```

- [ ] **Step 6: Run adversarial tests**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest tests/test_cli.py -k "adversarial_install_dir or adversarial_wrapper" -v`
Expected: 2 PASSED

- [ ] **Step 7: Commit**

```bash
git add src/auto_excel/config.py tests/test_cli.py
git commit -m "feat: add INSTALL_DIR and WRAPPER constants to config"
```

---

### Task 2: Refactor uninstall to use config constants

**Type:** Code task

**Context Brief:**
- **Files:** `src/auto_excel/cli.py` (modify — lines 76-98), `tests/test_cli.py` (read-only — existing uninstall tests)
- **Interface:** `uninstall()` function signature unchanged; only internal variable assignment changes
- **Dependencies:** Now reads `config.INSTALL_DIR` and `config.WRAPPER` instead of computing paths locally
- **Constraints:** Behavior must be identical — existing uninstall tests must pass without modification

**Files:**
- Modify: `src/auto_excel/cli.py:79-80`

- [ ] **Step 1: Run existing uninstall tests (baseline)**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest tests/test_cli.py -k "uninstall" -v`
Expected: All existing uninstall tests PASS

- [ ] **Step 2: Replace hardcoded paths with config constants**

In `src/auto_excel/cli.py`, inside the `uninstall()` function, replace lines 79-80:

Old:
```python
    install_dir = Path.home() / ".auto-excel"
    wrapper = Path.home() / ".local" / "bin" / "auto-excel"
```

New:
```python
    install_dir = config.INSTALL_DIR
    wrapper = config.WRAPPER
```

- [ ] **Step 3: Run uninstall tests to verify no regression**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest tests/test_cli.py -k "uninstall" -v`
Expected: All existing uninstall tests still PASS

- [ ] **Step 4: Run full test suite**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest -v`
Expected: All tests PASS

- [ ] **Step 5: Commit**

```bash
git add src/auto_excel/cli.py
git commit -m "refactor: use config constants in uninstall command"
```

---

### Task 3: Implement info command

**Type:** Code task

**Context Brief:**
- **Files:** `src/auto_excel/cli.py` (modify — add `info` command after `version`), `tests/test_cli.py` (modify — add info tests)
- **Interface:** `info() -> None` — Typer command, no parameters, prints 4 lines via `typer.echo`
- **Dependencies:** Reads `auto_excel.__version__`, `config.INSTALL_DIR`, `config.BASE_DIR`, `sys.version_info`
- **Constraints:** Output format must match spec: `auto-excel {version}` / `安装路径: {path}` / `Python:   {ver}` / `数据目录: {path}`. Uses `typer.echo` (consistent with existing `version` command).

**Files:**
- Modify: `src/auto_excel/cli.py` (add after `version` command, around line 30)
- Test: `tests/test_cli.py` (append new tests)

- [ ] **Step 1: Write failing tests**

Add to `tests/test_cli.py`:

```python
def test_info_shows_version():
    from auto_excel import __version__
    result = runner.invoke(app, ["info"])
    assert result.exit_code == 0
    assert __version__ in result.output

def test_info_shows_install_path():
    result = runner.invoke(app, ["info"])
    assert "安装路径" in result.output
    assert ".auto-excel" in result.output

def test_info_shows_python_version():
    import sys
    expected = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
    result = runner.invoke(app, ["info"])
    assert expected in result.output

def test_info_shows_data_dir():
    result = runner.invoke(app, ["info"])
    assert "数据目录" in result.output
    assert "marketing analysis" in result.output
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest tests/test_cli.py -k "test_info" -v`
Expected: FAIL — `info` command not found

- [ ] **Step 3: Implement info command**

In `src/auto_excel/cli.py`, add after the `version` command (after line 29):

```python
@app.command()
def info():
    """显示 auto-excel 环境信息。"""
    import sys
    from auto_excel import __version__
    typer.echo(f"auto-excel {__version__}")
    typer.echo(f"安装路径: {config.INSTALL_DIR}")
    typer.echo(f"Python:   {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}")
    typer.echo(f"数据目录: {config.BASE_DIR}")
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest tests/test_cli.py -k "test_info" -v`
Expected: 4 PASSED

- [ ] **Step 5: Write adversarial tests**

```python
def test_adversarial_info_exit_code_zero():
    result = runner.invoke(app, ["info"])
    assert result.exit_code == 0

def test_adversarial_info_has_four_lines():
    result = runner.invoke(app, ["info"])
    lines = [l for l in result.output.strip().split("\n") if l.strip()]
    assert len(lines) == 4

def test_adversarial_info_first_line_matches_version_cmd():
    r_ver = runner.invoke(app, ["version"])
    r_info = runner.invoke(app, ["info"])
    info_first_line = r_info.output.strip().split("\n")[0]
    assert info_first_line == r_ver.output.strip()
```

- [ ] **Step 6: Run adversarial tests**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest tests/test_cli.py -k "adversarial_info" -v`
Expected: 3 PASSED

- [ ] **Step 7: Commit**

```bash
git add src/auto_excel/cli.py tests/test_cli.py
git commit -m "feat: add info command showing version and environment details"
```

---

### Task 4: Add _parse_version_from_file helper

**Type:** Code task

**Context Brief:**
- **Files:** `src/auto_excel/cli.py` (modify — add helper function), `tests/test_cli.py` (modify — add tests)
- **Interface:** `_parse_version_from_file(path: Path) -> str` — reads a Python file and extracts `__version__` via regex. Returns version string or `"unknown"`.
- **Constraints:** Must handle double quotes, single quotes, and missing `__version__`. Regex: `r'__version__\s*=\s*["\']([^"\']+)["\']'`

**Files:**
- Modify: `src/auto_excel/cli.py` (add before `uninstall`, around line 64)
- Test: `tests/test_cli.py` (append new tests)

- [ ] **Step 1: Write failing tests**

```python
def test_parse_version_double_quotes(tmp_path):
    from auto_excel.cli import _parse_version_from_file
    f = tmp_path / "__init__.py"
    f.write_text('__version__ = "2.0.0"\n')
    assert _parse_version_from_file(f) == "2.0.0"

def test_parse_version_single_quotes(tmp_path):
    from auto_excel.cli import _parse_version_from_file
    f = tmp_path / "__init__.py"
    f.write_text("__version__ = '3.1.4'\n")
    assert _parse_version_from_file(f) == "3.1.4"

def test_parse_version_missing(tmp_path):
    from auto_excel.cli import _parse_version_from_file
    f = tmp_path / "__init__.py"
    f.write_text("# no version here\n")
    assert _parse_version_from_file(f) == "unknown"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest tests/test_cli.py -k "test_parse_version" -v`
Expected: FAIL — `_parse_version_from_file` not found

- [ ] **Step 3: Implement helper**

In `src/auto_excel/cli.py`, add `import re` to the top imports, then add before the `_remove_install_files` function:

```python
def _parse_version_from_file(path: Path) -> str:
    """从 Python 文件中提取 __version__ 值。"""
    text = path.read_text()
    match = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', text)
    return match.group(1) if match else "unknown"
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest tests/test_cli.py -k "test_parse_version" -v`
Expected: 3 PASSED

- [ ] **Step 5: Write adversarial tests**

```python
def test_adversarial_parse_version_extra_whitespace(tmp_path):
    from auto_excel.cli import _parse_version_from_file
    f = tmp_path / "__init__.py"
    f.write_text('__version__  =  "1.0.0"\n')
    assert _parse_version_from_file(f) == "1.0.0"

def test_adversarial_parse_version_with_surrounding_code(tmp_path):
    from auto_excel.cli import _parse_version_from_file
    f = tmp_path / "__init__.py"
    f.write_text('"""docstring"""\n\n__version__ = "4.5.6"\n\nother = 1\n')
    assert _parse_version_from_file(f) == "4.5.6"

def test_adversarial_parse_version_empty_file(tmp_path):
    from auto_excel.cli import _parse_version_from_file
    f = tmp_path / "__init__.py"
    f.write_text("")
    assert _parse_version_from_file(f) == "unknown"
```

- [ ] **Step 6: Run adversarial tests**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest tests/test_cli.py -k "adversarial_parse_version" -v`
Expected: 3 PASSED

- [ ] **Step 7: Commit**

```bash
git add src/auto_excel/cli.py tests/test_cli.py
git commit -m "feat: add _parse_version_from_file helper for upgrade command"
```

---

### Task 5: Implement upgrade — install dir missing & already latest paths

**Type:** Code task

**Context Brief:**
- **Files:** `src/auto_excel/cli.py` (modify — add `upgrade` command), `tests/test_cli.py` (modify — add tests)
- **Interface:** `upgrade() -> None` — Typer command, no parameters. Calls `subprocess.run` for git commands. Uses `config.INSTALL_DIR`. Raises `typer.Exit(1)` on error.
- **Dependencies:** Reads `config.INSTALL_DIR`, `auto_excel.__version__`. Calls `subprocess.run` with git args.
- **Constraints:** Must add `import subprocess` to cli.py top-level imports. Mock pattern: `monkeypatch.setattr(subprocess, "run", mock_fn)` + `monkeypatch.setattr(cfg, "INSTALL_DIR", tmp_path)`.

**Files:**
- Modify: `src/auto_excel/cli.py` (add `import subprocess` at top; add `upgrade` command after `info`)
- Test: `tests/test_cli.py`

- [ ] **Step 1: Write failing tests for both early-exit paths**

```python
def test_upgrade_no_install_dir(tmp_path, monkeypatch):
    import auto_excel.config as cfg
    monkeypatch.setattr(cfg, "INSTALL_DIR", tmp_path / "nonexistent")
    result = runner.invoke(app, ["upgrade"])
    assert result.exit_code == 1
    assert "未检测到" in result.output

def test_upgrade_already_latest(tmp_path, monkeypatch):
    import auto_excel.config as cfg
    import subprocess
    monkeypatch.setattr(cfg, "INSTALL_DIR", tmp_path)
    (tmp_path / ".git").mkdir()
    def mock_run(args, **kwargs):
        if "fetch" in args:
            return subprocess.CompletedProcess(args, 0, "", "")
        if "rev-parse" in args:
            return subprocess.CompletedProcess(args, 0, "abc123\n", "")
        return subprocess.CompletedProcess(args, 0, "", "")
    monkeypatch.setattr(subprocess, "run", mock_run)
    result = runner.invoke(app, ["upgrade"])
    assert result.exit_code == 0
    assert "已是最新" in result.output
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest tests/test_cli.py::test_upgrade_no_install_dir tests/test_cli.py::test_upgrade_already_latest -v`
Expected: FAIL — `upgrade` command not found

- [ ] **Step 3: Implement upgrade command (early-exit paths only)**

Add `import subprocess` to cli.py imports. Add the `upgrade` command after `info`:

```python
@app.command()
def upgrade():
    """从 GitHub 升级到最新版本。"""
    from auto_excel import __version__
    install_dir = config.INSTALL_DIR
    if not (install_dir / ".git").is_dir():
        typer.echo("错误：未检测到安装目录，请重新安装：")
        typer.echo("  curl -sSL https://raw.githubusercontent.com/SuperAllenYao/auto-excel/master/install.sh | bash")
        raise typer.Exit(1)
    typer.echo(f"当前版本: v{__version__}")
    result = subprocess.run(
        ["git", "-C", str(install_dir), "fetch", "origin", "master"],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        typer.echo("错误：无法连接到远程仓库，请检查网络")
        typer.echo(result.stderr.strip())
        raise typer.Exit(1)
    local = subprocess.run(
        ["git", "-C", str(install_dir), "rev-parse", "HEAD"],
        capture_output=True, text=True,
    )
    remote = subprocess.run(
        ["git", "-C", str(install_dir), "rev-parse", "origin/master"],
        capture_output=True, text=True,
    )
    if local.stdout.strip() == remote.stdout.strip():
        typer.echo(f"已是最新版本 v{__version__}，无需升级")
        return
    # Pull + sync will be implemented in next task
    typer.echo("有新版本可用，升级中...")
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest tests/test_cli.py::test_upgrade_no_install_dir tests/test_cli.py::test_upgrade_already_latest -v`
Expected: 2 PASSED

- [ ] **Step 5: Write adversarial tests**

```python
def test_adversarial_upgrade_shows_reinstall_cmd(tmp_path, monkeypatch):
    import auto_excel.config as cfg
    monkeypatch.setattr(cfg, "INSTALL_DIR", tmp_path / "gone")
    result = runner.invoke(app, ["upgrade"])
    assert "curl" in result.output
    assert "install.sh" in result.output

def test_adversarial_upgrade_fetch_network_error(tmp_path, monkeypatch):
    import auto_excel.config as cfg
    import subprocess
    monkeypatch.setattr(cfg, "INSTALL_DIR", tmp_path)
    (tmp_path / ".git").mkdir()
    def mock_run(args, **kwargs):
        if "fetch" in args:
            return subprocess.CompletedProcess(args, 1, "", "fatal: unable to access")
        return subprocess.CompletedProcess(args, 0, "", "")
    monkeypatch.setattr(subprocess, "run", mock_run)
    result = runner.invoke(app, ["upgrade"])
    assert result.exit_code == 1
    assert "网络" in result.output or "远程仓库" in result.output
```

- [ ] **Step 6: Run adversarial tests**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest tests/test_cli.py -k "adversarial_upgrade" -v`
Expected: 2 PASSED

- [ ] **Step 7: Commit**

```bash
git add src/auto_excel/cli.py tests/test_cli.py
git commit -m "feat: add upgrade command with install-dir check and already-latest detection"
```

---

### Task 6: Complete upgrade — successful upgrade and error paths

**Type:** Code task

**Context Brief:**
- **Files:** `src/auto_excel/cli.py` (modify — complete `upgrade` command), `tests/test_cli.py` (modify — add tests)
- **Interface:** Complete the `upgrade()` function: after detecting a new version is available, run `git pull` + `uv sync`, then parse and display new version.
- **Dependencies:** Uses `_parse_version_from_file()` from Task 4. Calls `subprocess.run` with git pull and uv sync args.
- **Constraints:** `git pull` args: `["git", "-C", str(install_dir), "pull", "origin", "master"]`. `uv sync` args: `["uv", "sync"]` with `cwd=str(install_dir)`. New version is read from `install_dir / "src" / "auto_excel" / "__init__.py"`.

**Files:**
- Modify: `src/auto_excel/cli.py` (replace placeholder at end of `upgrade`)
- Test: `tests/test_cli.py`

- [ ] **Step 1: Write failing test for successful upgrade**

```python
def test_upgrade_success(tmp_path, monkeypatch):
    import auto_excel.config as cfg
    import subprocess
    monkeypatch.setattr(cfg, "INSTALL_DIR", tmp_path)
    (tmp_path / ".git").mkdir()
    init_dir = tmp_path / "src" / "auto_excel"
    init_dir.mkdir(parents=True)
    (init_dir / "__init__.py").write_text('__version__ = "2.0.0"\n')
    def mock_run(args, **kwargs):
        if "fetch" in args:
            return subprocess.CompletedProcess(args, 0, "", "")
        if "rev-parse" in args and "origin/master" in args:
            return subprocess.CompletedProcess(args, 0, "def456\n", "")
        if "rev-parse" in args:
            return subprocess.CompletedProcess(args, 0, "abc123\n", "")
        if "pull" in args:
            return subprocess.CompletedProcess(args, 0, "", "")
        if args and args[0] == "uv":
            return subprocess.CompletedProcess(args, 0, "", "")
        return subprocess.CompletedProcess(args, 0, "", "")
    monkeypatch.setattr(subprocess, "run", mock_run)
    result = runner.invoke(app, ["upgrade"])
    assert result.exit_code == 0
    assert "升级成功" in result.output
    assert "2.0.0" in result.output
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest tests/test_cli.py::test_upgrade_success -v`
Expected: FAIL — "升级成功" not in output (placeholder code doesn't complete the flow)

- [ ] **Step 3: Complete upgrade implementation**

In `src/auto_excel/cli.py`, replace the placeholder at the end of `upgrade()` (the `typer.echo("有新版本可用，升级中...")` line) with the full pull + sync + version display logic:

```python
    # After the rev-parse comparison determines update is needed:
    typer.echo("检测到新版本，正在升级...")
    result = subprocess.run(
        ["git", "-C", str(install_dir), "pull", "origin", "master"],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        typer.echo("错误：拉取更新失败")
        typer.echo(result.stderr.strip())
        raise typer.Exit(1)
    result = subprocess.run(
        ["uv", "sync"], capture_output=True, text=True, cwd=str(install_dir),
    )
    if result.returncode != 0:
        typer.echo("错误：依赖同步失败")
        typer.echo(result.stderr.strip())
        raise typer.Exit(1)
    new_version = _parse_version_from_file(
        install_dir / "src" / "auto_excel" / "__init__.py"
    )
    typer.echo(f"升级成功: v{__version__} → v{new_version}")
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest tests/test_cli.py -k "upgrade" -v`
Expected: All upgrade tests PASS (including Task 5 tests)

- [ ] **Step 5: Write adversarial tests for error paths**

```python
def test_adversarial_upgrade_pull_fails(tmp_path, monkeypatch):
    import auto_excel.config as cfg
    import subprocess
    monkeypatch.setattr(cfg, "INSTALL_DIR", tmp_path)
    (tmp_path / ".git").mkdir()
    def mock_run(args, **kwargs):
        if "fetch" in args:
            return subprocess.CompletedProcess(args, 0, "", "")
        if "rev-parse" in args and "origin/master" in args:
            return subprocess.CompletedProcess(args, 0, "def456\n", "")
        if "rev-parse" in args:
            return subprocess.CompletedProcess(args, 0, "abc123\n", "")
        if "pull" in args:
            return subprocess.CompletedProcess(args, 1, "", "merge conflict")
        return subprocess.CompletedProcess(args, 0, "", "")
    monkeypatch.setattr(subprocess, "run", mock_run)
    result = runner.invoke(app, ["upgrade"])
    assert result.exit_code == 1
    assert "拉取更新失败" in result.output

def test_adversarial_upgrade_uv_sync_fails(tmp_path, monkeypatch):
    import auto_excel.config as cfg
    import subprocess
    monkeypatch.setattr(cfg, "INSTALL_DIR", tmp_path)
    (tmp_path / ".git").mkdir()
    def mock_run(args, **kwargs):
        if "fetch" in args:
            return subprocess.CompletedProcess(args, 0, "", "")
        if "rev-parse" in args and "origin/master" in args:
            return subprocess.CompletedProcess(args, 0, "def456\n", "")
        if "rev-parse" in args:
            return subprocess.CompletedProcess(args, 0, "abc123\n", "")
        if "pull" in args:
            return subprocess.CompletedProcess(args, 0, "", "")
        if args and args[0] == "uv":
            return subprocess.CompletedProcess(args, 1, "", "error: no such option")
        return subprocess.CompletedProcess(args, 0, "", "")
    monkeypatch.setattr(subprocess, "run", mock_run)
    result = runner.invoke(app, ["upgrade"])
    assert result.exit_code == 1
    assert "依赖同步失败" in result.output

def test_adversarial_upgrade_shows_old_and_new_version(tmp_path, monkeypatch):
    """Upgrade output must contain both old and new version strings."""
    import auto_excel.config as cfg
    import subprocess
    from auto_excel import __version__
    monkeypatch.setattr(cfg, "INSTALL_DIR", tmp_path)
    (tmp_path / ".git").mkdir()
    init_dir = tmp_path / "src" / "auto_excel"
    init_dir.mkdir(parents=True)
    (init_dir / "__init__.py").write_text('__version__ = "9.9.9"\n')
    def mock_run(args, **kwargs):
        if "fetch" in args:
            return subprocess.CompletedProcess(args, 0, "", "")
        if "rev-parse" in args and "origin/master" in args:
            return subprocess.CompletedProcess(args, 0, "new\n", "")
        if "rev-parse" in args:
            return subprocess.CompletedProcess(args, 0, "old\n", "")
        return subprocess.CompletedProcess(args, 0, "", "")
    monkeypatch.setattr(subprocess, "run", mock_run)
    result = runner.invoke(app, ["upgrade"])
    assert __version__ in result.output
    assert "9.9.9" in result.output
```

- [ ] **Step 6: Run adversarial tests**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest tests/test_cli.py -k "adversarial_upgrade" -v`
Expected: All adversarial upgrade tests PASS

- [ ] **Step 7: Run full test suite**

Run: `cd /Users/allenyao/work_project/auto-excel && uv run pytest -v`
Expected: All tests PASS

- [ ] **Step 8: Commit**

```bash
git add src/auto_excel/cli.py tests/test_cli.py
git commit -m "feat: complete upgrade command with pull, sync, and error handling"
```
