#!/usr/bin/env python3
"""
Offline-friendly XLSX -> CSV converter.

- Reads an .xlsx file (all sheets).
- Exports each sheet to a separate UTF-8 (with BOM) comma-separated CSV.
- Avoids pandas/numpy to keep the exe lightweight.

Example:
  xlsx_to_csv.exe --input "your.xlsx" --out-dir ".\\out"
"""

from __future__ import annotations

import argparse
import csv
import re
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook


def _safe_name(name: str) -> str:
    s = str(name or "").strip()
    s = re.sub(r"[^\w\u4e00-\u9fff.-]+", "_", s)
    return (s[:120] if s else "Sheet").strip("_") or "Sheet"


def _cell_to_text(v: object) -> str:
    if v is None:
        return ""
    if isinstance(v, str):
        return v
    return str(v)


def convert_one(xlsx: Path, out_dir: Path) -> list[Path]:
    if not xlsx.is_file():
        raise SystemExit(f"输入文件不存在: {xlsx}")
    if xlsx.suffix.lower() not in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        raise SystemExit(f"仅支持 xlsx/xlsm/xltx/xltm: {xlsx}")
    out_dir.mkdir(parents=True, exist_ok=True)

    written: list[Path] = []
    base = _safe_name(xlsx.stem)

    wb = load_workbook(filename=xlsx, read_only=True, data_only=False)
    try:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            out = out_dir / f"{base}__{_safe_name(sheet_name)}.csv"
            with out.open("w", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(
                    f,
                    delimiter=",",
                    quotechar='"',
                    quoting=csv.QUOTE_MINIMAL,
                    lineterminator="\n",
                )
                for row in ws.iter_rows(values_only=True):
                    w.writerow([_cell_to_text(v) for v in row])
            written.append(out)
    finally:
        try:
            wb.close()
        except Exception:
            pass
    return written


def _write_run_info(*, xlsx: Path, out_dir: Path, written: list[Path]) -> Path:
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    info_path = out_dir / f"{_safe_name(xlsx.stem)}__conversion_info__{ts}.txt"
    lines: list[str] = [
        "xlsx_to_csv conversion info",
        f"time: {datetime.now().isoformat(timespec='seconds')}",
        f"input: {xlsx}",
        f"output_dir: {out_dir}",
        f"csv_count: {len(written)}",
        "",
        "files:",
        *[f"- {p.name}" for p in written],
        "",
    ]
    info_path.write_text("\n".join(lines), encoding="utf-8")
    return info_path


def _pick_file_gui() -> Path | None:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except Exception:
        return None

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        p = filedialog.askopenfilename(
            title="选择要转换的 XLSX 文件",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm *.xltx *.xltm"),
                ("All files", "*.*"),
            ],
        )
        if not p:
            return None
        xlsx = Path(p).expanduser().resolve()
        if not xlsx.is_file():
            messagebox.showerror("错误", f"文件不存在：{xlsx}")
            return None
        return xlsx
    finally:
        try:
            root.destroy()
        except Exception:
            pass


def main() -> int:
    ap = argparse.ArgumentParser(description="XLSX -> UTF-8 comma CSV (per sheet)")
    ap.add_argument("--input", required=False, help="Path to .xlsx file")
    ap.add_argument(
        "--out-dir",
        required=False,
        help="Output directory for CSV files (default: same directory as input file)",
    )
    ap.add_argument(
        "--gui",
        action="store_true",
        help="Force GUI file picker when available",
    )
    args = ap.parse_args()

    xlsx: Path | None = None
    if args.input and not args.gui:
        xlsx = Path(args.input).expanduser().resolve()
    else:
        xlsx = _pick_file_gui()

    if xlsx is None:
        raise SystemExit(
            "未选择输入文件。命令行用法示例：xlsx_to_csv.exe --input your.xlsx --out-dir .\\out"
        )

    out_dir = (
        Path(args.out_dir).expanduser().resolve()
        if args.out_dir
        else xlsx.parent.resolve()
    )
    paths = convert_one(xlsx, out_dir)
    info_path = _write_run_info(xlsx=xlsx, out_dir=out_dir, written=paths)

    print(f"OK: wrote {len(paths)} CSV file(s) to {out_dir}")
    for p in paths:
        print(f" - {p.name}")
    print(f"info: {info_path.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

