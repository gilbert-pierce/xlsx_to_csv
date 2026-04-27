## xlsx_to_csv_offline_tool

把 Excel `.xlsx` 稳定转换成 **UTF-8（带 BOM）+ 逗号分隔** 的 CSV（每个 sheet 一个 CSV）。

### 运行（Python）

```bash
python xlsx_to_csv.py --input "你的文件.xlsx" --out-dir "./out"
```

也支持对文件夹内所有 xlsx 批量转换（可选递归）：

```bash
python xlsx_to_csv.py --input-dir "./excels" --recursive
```

### 打包 Win10 x64 离线 exe（PyInstaller）

> 结论：请在 **Windows 10 x64** 环境打包（避免跨平台二进制依赖问题）。

1) 在“有网”的机器下载 wheels（拷贝到离线机）：

```bash
python -m pip download -d wheels pyinstaller openpyxl et-xmlfile
```

2) 在 Win10 离线机安装依赖：

```bat
py -m venv venv
venv\Scripts\pip install --no-index --find-links wheels pyinstaller openpyxl
```

3) 打包：

```bat
venv\Scripts\pyinstaller --onefile --name xlsx_to_csv_cli xlsx_to_csv.py
venv\Scripts\pyinstaller --onefile --noconsole --name xlsx_to_csv_gui xlsx_to_csv.py
```

输出：`dist\xlsx_to_csv_cli.exe`、`dist\xlsx_to_csv_gui.exe`

4) 运行：

```bat
dist\xlsx_to_csv_cli.exe --input "C:\path\file.xlsx" --out-dir "C:\path\out"
```

### 运行（exe，带交互界面）

- **直接双击** `dist\xlsx_to_csv_gui.exe`：会弹出文件选择框；转换结果默认输出到**输入文件同目录**，并生成一份 `*_conversion_info_*.txt` 记录本次转换信息。
- **命令行**：仍可用 `--input/--out-dir` 指定路径（适合批处理/自动化）。

也支持文件夹批量：

```bat
dist\xlsx_to_csv_cli.exe --input-dir "C:\path\excels" --recursive
```

