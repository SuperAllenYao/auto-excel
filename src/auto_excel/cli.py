"""CLI entry point."""
import logging
import shutil
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
