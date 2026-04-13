"""Rich-based terminal output module for auto-excel."""
from __future__ import annotations

from rich.console import Console
from rich.table import Table

_console = Console()


def print_start(count: int, console: Console | None = None) -> None:
    """打印启动消息，显示待处理文件数量。"""
    con = console or _console
    con.print(f"🚀 开始处理，共 [bold]{count}[/bold] 个文件")


def print_file_start(idx: int, total: int, name: str, console: Console | None = None) -> None:
    """打印当前处理的文件信息。"""
    con = console or _console
    con.print(f"\n📄 处理文件 [{idx}/{total}]: [bold cyan]{name}[/bold cyan]")


def print_step(msg: str, console: Console | None = None) -> None:
    """打印处理步骤消息。"""
    con = console or _console
    con.print(f"  ⚙️  {msg}")


def print_file_done(duration: float, console: Console | None = None) -> None:
    """打印文件处理完成消息，显示耗时。"""
    con = console or _console
    con.print(f"  ✅ 完成，耗时 [green]{duration:.2f}[/green] 秒")


def print_file_error(name: str, err: Exception | str, console: Console | None = None) -> None:
    """打印文件处理错误消息。"""
    con = console or _console
    con.print(f"  ❌ 处理 [bold]{name}[/bold] 时出错：[red]{err}[/red]")


def print_report(results: list[dict], console: Console | None = None) -> None:
    """打印处理结果汇总表格。

    Args:
        results: 包含 filename、status、duration 字段的字典列表。
                 status 为 'success' 或 'error'。
    """
    con = console or _console

    table = Table(title="📊 处理报告", show_header=True, header_style="bold magenta")
    table.add_column("文件名", style="cyan", no_wrap=True)
    table.add_column("状态", justify="center")
    table.add_column("耗时", justify="right")

    for item in results:
        filename = item.get("filename", "")
        status = item.get("status", "")
        duration = item.get("duration", 0)

        if status == "success":
            status_display = "✅ 成功"
            duration_display = f"{duration:.2f}s"
        else:
            status_display = "❌ 失败"
            duration_display = "-"

        table.add_row(filename, status_display, duration_display)

    con.print(table)


def print_no_files(console: Console | None = None) -> None:
    """打印无文件可处理的消息。"""
    con = console or _console
    con.print("⚠️  没有找到需要处理的文件，请检查输入路径。")


def print_exit(console: Console | None = None) -> None:
    """打印退出告别消息。"""
    con = console or _console
    con.print("👋 感谢使用 auto-excel，再见！")
