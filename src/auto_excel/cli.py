"""CLI entry point."""
import logging
import re
import shutil
import subprocess
import time
from datetime import date
from pathlib import Path

import typer

from auto_excel import config, display, state
from auto_excel.processor import process_file

app = typer.Typer()


def _setup_logging():
    config.LOG_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        filename=str(config.LOG_DIR / f"{date.today()}.log"),
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )


@app.command()
def version():
    """Show version information."""
    from auto_excel import __version__
    typer.echo(f"auto-excel {__version__}")


@app.command()
def info():
    """显示 auto-excel 环境信息。"""
    import sys
    from auto_excel import __version__
    typer.echo(f"auto-excel {__version__}")
    typer.echo(f"安装路径: {config.INSTALL_DIR}")
    typer.echo(f"Python:   {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}")
    typer.echo(f"数据目录: {config.BASE_DIR}")


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


@app.command()
def on():
    """Process all unprocessed Excel files in Raw/."""
    _setup_logging()
    config.RAW_DIR.mkdir(parents=True, exist_ok=True)
    config.NEW_DIR.mkdir(parents=True, exist_ok=True)
    all_files = sorted(config.RAW_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_ctime)
    pending = [f for f in all_files if not state.is_processed(f.name)]
    if not pending:
        display.print_no_files()
        display.print_exit()
        return
    display.print_start(len(pending))
    results = []
    for idx, src in enumerate(pending, start=1):
        dst = config.NEW_DIR / src.name
        display.print_file_start(idx, len(pending), src.name)
        t0 = time.monotonic()
        try:
            process_file(src, dst, on_step=display.print_step)
            state.save_file_state(src.name)
            dur = time.monotonic() - t0
            display.print_file_done(dur)
            results.append({"filename": src.name, "status": "success", "duration": dur})
        except Exception as exc:
            logging.exception("处理 %s 时出错", src.name)
            display.print_file_error(src.name, exc)
            results.append({"filename": src.name, "status": "error", "duration": 0})
    display.print_report(results)
    display.print_exit()


def _parse_version_from_file(path: Path) -> str:
    """从 Python 文件中提取 __version__ 值。"""
    text = path.read_text()
    match = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', text)
    return match.group(1) if match else "unknown"


def _remove_install_files(install_dir: Path, wrapper: Path) -> list[str]:
    """删除安装目录和 wrapper 脚本，返回实际删除的路径列表。"""
    removed = []
    if wrapper.exists():
        wrapper.unlink()
        removed.append(str(wrapper))
    if install_dir.exists():
        shutil.rmtree(install_dir)
        removed.append(str(install_dir))
    return removed


@app.command()
def uninstall():
    """完全卸载 auto-excel（保留已处理的文档）。"""
    install_dir = config.INSTALL_DIR
    wrapper = config.WRAPPER

    typer.echo("以下文件将被删除：")
    typer.echo(f"  {install_dir}")
    typer.echo(f"  {wrapper}")
    typer.echo("以下内容不会被删除：")
    typer.echo(f"  ~/Desktop/marketing analysis/（已处理文档完整保留）")
    typer.echo("")

    typer.confirm("确认卸载？", abort=True)

    removed = _remove_install_files(install_dir, wrapper)

    if removed:
        for path in removed:
            typer.echo(f"已删除：{path}")
    else:
        typer.echo("未找到需要删除的文件，可能已经卸载。")
    typer.echo("auto-excel 已卸载。")
