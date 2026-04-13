"""CLI entry point."""
import logging
import time
from datetime import date

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
    typer.echo("auto-excel 0.1.0")


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
