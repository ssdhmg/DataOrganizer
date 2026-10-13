# -*- coding: utf-8 -*-
"""
自动整理数据脚本
功能：
1. 选择多个原始数据 CSV 文件
2. 选择模板 xlsx
3. 选择寄存器截图 xlsx（图片从左到右排列）
4. 选择输出文件夹后自动处理

性能说明（Win7 / 老硬件优化）：
- Excel 进程只启动一次，整个处理期间保持运行，最后统一 Quit，不再反复创建/销毁。
- 模板工作簿只打开一次作为母版，之后每个 CSV 用 Worksheets.Copy 复制一份副本，
  避免重复解析模板、重复加载图表与形状（这是老机器上最耗时的部分）。
- 写入数据用二维数组一次性赋值（单次 COM 往返），不再逐单元格调用。
- 图表数据校验从"先落盘再 reopen"改为直接读取母版/副本的图表系列引用，
  省掉一次 SaveAs + 一次 Open。
"""

import sys
import os
import re
import csv
import shutil
import tempfile
import threading
from pathlib import Path
from datetime import datetime


def get_base_dir():
    """获取程序运行目录（打包后为 exe 所在目录）"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from PIL import Image as PILImage

try:
    import openpyxl
    import openpyxl.utils
except ImportError:
    openpyxl = None

try:
    import pythoncom
    import win32com.client

    COM_AVAILABLE = True
except ImportError:
    pythoncom = None
    win32com = None
    COM_AVAILABLE = False


# ------------------------------------------------------------
# 常量：校验相关
# ------------------------------------------------------------
# 图表系列设定值允许的最大偏差（百分比）
CURRENT_TOLERANCE_PCT = 10.0
# 参与电流平均值计算的"前几行"
CURRENT_SAMPLE_ROWS = 5

# 图表里"Offset"系列所属的参考块（无对应负载电流，不参与校验）
REFERENCE_BLOCK_OFFSET_COL = 7   # 第1个 Offset 列 G，属于 Vref 参考块

# 需要校验的图表对象序号（图表1 是 Vref 参考块，图表2~8 是带电流设定值的图）
CHECK_CHART_INDEXES = (2, 3, 4, 5, 6, 7, 8)


# ------------------------------------------------------------
# 辅助函数
# ------------------------------------------------------------
def read_csv_rows(csv_path):
    """读取 CSV 所有行，自动尝试常见编码。"""
    encodings = ["utf-8-sig", "utf-8", "gbk", "gb18030"]
    last_err = None
    for enc in encodings:
        try:
            with open(csv_path, newline="", encoding=enc) as f:
                reader = csv.reader(f)
                return list(reader)
        except UnicodeDecodeError as e:
            last_err = e
            continue
    raise last_err


def get_csv_data_and_id(csv_path, max_rows=None):
    """
    从 CSV 中获取：
    1. 第 2 行到最后一行的数据（1-based 行号），若指定 max_rows，则只取前 max_rows 行数据
    2. ID 列的值（优先查找表头为 ID 的列，否则使用第一列）
    """
    rows = read_csv_rows(csv_path)
    if not rows:
        return [], ""

    header = rows[0]
    id_col = None
    for i, cell in enumerate(header):
        if str(cell).strip().lower() == "id":
            id_col = i
            break
    if id_col is None:
        id_col = 0

    # 修改：根据 max_rows 截取数据行
    if max_rows is not None:
        data_rows = rows[1:1 + max_rows]
    else:
        data_rows = rows[1:]

    id_value = ""
    for r in data_rows:
        if id_col < len(r) and str(r[id_col]).strip():
            id_value = str(r[id_col]).strip()
            break

    return data_rows, id_value


def to_number(value):
    """尽量把单元格值转成 float，失败返回 None。"""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    # 兼容 Excel 把数字当文本存的情况，以及 "1.23E+02" 之类
    try:
        return float(text)
    except ValueError:
        m = re.search(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", text)
        if m:
            try:
                return float(m.group(0))
            except ValueError:
                return None
    return None


def sanitize_filename_part(value):
    """清理文件名中的非法字符，替换为下划线"""
    return re.sub(r'[<>:"/\\|?*#%~\s]', '_', value)


def get_id_from_excel_file(file_path):
    """从保存的 xlsx 文件中读取 ID 列第一个非空值（全表数据行）"""
    if not openpyxl:
        return None
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        ws = wb.active
        # 查找表头中的 ID 列
        id_col = None
        for cell in ws[1]:
            if cell.value and str(cell.value).strip().lower() == "id":
                id_col = cell.column
                break
        if id_col is None:
            id_col = 1  # 默认第一列
        # 从第2行到最后一行查找第一个非空 ID
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=id_col, max_col=id_col, values_only=True):
            if row and row[0] and str(row[0]).strip():
                wb.close()
                return str(row[0]).strip()
        wb.close()
        return None
    except Exception:
        return None


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("数据自动整理工具")
        self.root.geometry("820x700")  # 稍微增加高度以容纳新输入框

        self.raw_files = []
        self.template_path = None
        self.reg_path = None
        self.output_dir = None
        self.temp_dir = None

        self.raw_var = tk.StringVar(value="未选择文件")
        self.template_var = tk.StringVar(value="未选择模板")
        self.reg_var = tk.StringVar(value="未选择截图文件")
        self.output_var = tk.StringVar(value="未选择输出文件夹")
        # 新增：文件名后缀与数据行数输入变量
        self.suffix_var = tk.StringVar(value="(-40_125)")
        self.row_count_var = tk.StringVar(value="80")

        self.build_ui()

    # ---------------------- UI ----------------------
    def build_ui(self):
        # 第一行：原始数据
        tk.Button(self.root, text="选择原始数据(可多选)", width=22,
                  command=self.choose_raw).grid(row=0, column=0, padx=10, pady=8, sticky="w")
        tk.Label(self.root, textvariable=self.raw_var, anchor="w", fg="blue",
                 width=70).grid(row=0, column=1, padx=10, pady=8, sticky="w")

        # 第二行：模板文件
        tk.Button(self.root, text="选择模板文件", width=22,
                  command=self.choose_template).grid(row=1, column=0, padx=10, pady=8, sticky="w")
        tk.Label(self.root, textvariable=self.template_var, anchor="w", fg="blue",
                 width=70).grid(row=1, column=1, padx=10, pady=8, sticky="w")

        # 第三行：寄存器截图
        tk.Button(self.root, text="选择寄存器截图文件", width=22,
                  command=self.choose_reg).grid(row=2, column=0, padx=10, pady=8, sticky="w")
        tk.Label(self.root, textvariable=self.reg_var, anchor="w", fg="blue",
                 width=70).grid(row=2, column=1, padx=10, pady=8, sticky="w")

        # 新增：第四行 - 文件名后缀输入
        tk.Label(self.root, text="文件名后缀:", width=22, anchor="w").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        tk.Entry(self.root, textvariable=self.suffix_var, width=40).grid(row=3, column=1, padx=10, pady=5, sticky="w")

        # 新增：第五行 - 迁移数据行数输入
        tk.Label(self.root, text="迁移数据行数:", width=22, anchor="w").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        tk.Entry(self.root, textvariable=self.row_count_var, width=40).grid(row=4, column=1, padx=10, pady=5, sticky="w")

        # 第六行：输出文件夹
        tk.Button(self.root, text="选择输出文件夹并开始", width=22,
                  command=self.choose_output_and_process).grid(row=5, column=0, padx=10, pady=15, sticky="w")
        tk.Label(self.root, textvariable=self.output_var, anchor="w", fg="blue",
                 width=70).grid(row=5, column=1, padx=10, pady=15, sticky="w")

        # 日志区域
        self.log_text = scrolledtext.ScrolledText(self.root, height=16, state="disabled",
                                                  font=("Consolas", 9))
        self.log_text.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(6, weight=1)

    def choose_raw(self):
        files = filedialog.askopenfilenames(
            title="选择原始数据 CSV（可多选）",
            filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")]
        )
        if files:
            self.raw_files = list(files)
            self.raw_var.set(f"已选择 {len(files)} 个文件")
            for f in self.raw_files:
                self.log(f"原始数据: {f}")

    def choose_template(self):
        path = filedialog.askopenfilename(
            title="选择模板文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if path:
            self.template_path = path
            self.template_var.set(path)
            self.log(f"模板文件: {path}")

    def choose_reg(self):
        path = filedialog.askopenfilename(
            title="选择寄存器截图文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if path:
            self.reg_path = path
            self.reg_var.set(path)
            self.log(f"寄存器截图文件: {path}")

    def choose_output_and_process(self):
        if not self.raw_files:
            messagebox.showwarning("提示", "请先选择原始数据 CSV 文件")
            return
        if not self.template_path:
            messagebox.showwarning("提示", "请先选择模板文件")
            return
        if not self.reg_path:
            messagebox.showwarning("提示", "请先选择寄存器截图文件")
            return

        output_dir = filedialog.askdirectory(title="选择存放处理结果的文件夹")
        if not output_dir:
            return

        self.output_dir = os.path.abspath(os.path.normpath(output_dir))
        self.output_var.set(output_dir)
        self.log(f"输出文件夹: {output_dir}")

        # 在后台线程执行，避免界面卡死
        threading.Thread(target=self.process_work, daemon=True).start()

    # ---------------------- 日志 ----------------------
    def log(self, msg):
        def _append():
            self.log_text.configure(state="normal")
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
            self.log_text.configure(state="disabled")

        if threading.current_thread() is threading.main_thread():
            _append()
        else:
            self.root.after(0, _append)

    # ---------------------- 图片提取 ----------------------
    def extract_images_from_xlsx(self, xlsx_path):
        """从寄存器截图 xlsx 中提取所有图片，并按从左到右排列。"""
        if not openpyxl:
            self.log("请先安装 openpyxl：pip install openpyxl")
            return []

        self.temp_dir = tempfile.mkdtemp(prefix="reg_imgs_")
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)

        items = []
        for ws in wb.worksheets:
            for img in ws._images:
                try:
                    col = img.anchor._from.col
                    row = img.anchor._from.row
                except Exception:
                    col, row = 9999, 9999
                items.append((ws.title, col, row, img))

        wb.close()

        # 按 sheet、列、行排序，即图片从左到右
        items.sort(key=lambda x: (x[0], x[1], x[2]))

        paths = []
        for idx, (sheet, col, row, img) in enumerate(items):
            try:
                data = img._data()
            except Exception:
                data = None

            if data:
                ext = getattr(img, "format", "png") or "png"
                if ext.lower() not in ["png", "jpg", "jpeg", "gif", "bmp", "tiff"]:
                    ext = "png"
                img_path = os.path.join(self.temp_dir, f"reg_{idx}_{col}_{row}.{ext}")
                with open(img_path, "wb") as f:
                    f.write(data)
                paths.append(img_path)

        self.log(f"从截图文件中提取到 {len(paths)} 张图片")
        return paths

    # ---------------------- Excel 操作 ----------------------
    def set_chart8_title(self, ws, id_value):
        """修改图表 8 标题，把 # 前面的数字改成 ID 值。"""
        try:
            chart_obj = ws.ChartObjects(8)
            chart = chart_obj.Chart

            try:
                chart.HasTitle = True
            except Exception:
                pass

            old_title = ""
            try:
                old_title = chart.ChartTitle.Text or ""
            except Exception:
                old_title = ""

            if old_title:
                # 只替换 # 前面的数字
                new_title = re.sub(r"\d+(?=\s*#)", str(id_value), old_title, count=1)
                if not re.search(r"\d+\s*#", old_title) and "#" in old_title:
                    new_title = old_title.replace("#", f"{id_value}#", 1)
            else:
                new_title = f"{id_value}#"

            chart.ChartTitle.Text = new_title
            self.log(f"图表 8 标题已修改为: {new_title}")
        except Exception as e:
            self.log(f"修改图表 8 标题失败: {e}")

    def update_textbox9_time(self, ws):
        """
        更新文本框 9 里的测试时间。
        只替换"测试时间："后面那一行的内容，避免把后面几行（测试人员/地点等）一起吃掉。
        """
        now_str = datetime.now().strftime("%Y.%m.%d")
        updated = False

        try:
            for i in range(ws.Shapes.Count, 0, -1):
                shape = ws.Shapes(i)
                try:
                    if shape.Type == 17:  # msoTextBox
                        name = shape.Name or ""
                        if re.search(r"(TextBox\s*9|文本框\s*9)", name, re.IGNORECASE):
                            old_text = ""
                            try:
                                old_text = shape.TextFrame2.TextRange.Text
                            except Exception:
                                try:
                                    old_text = shape.TextFrame.Characters().Text
                                except Exception:
                                    continue

                            new_text = old_text
                            # 优先按"日期样式"替换，最稳妥
                            date_pat = r"\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}"
                            if re.search(r"测试时间[:：]\s*" + date_pat, new_text):
                                new_text = re.sub(
                                    r"(测试时间[:：]\s*)" + date_pat,
                                    r"\g<1>" + now_str,
                                    new_text,
                                    count=1
                                )
                            else:
                                # 没有现成日期时，只替换冒号后到行尾的内容
                                # [^\r\n]* 且不加 \s*（\s 会跨行，容易把后续行吞掉）
                                new_text = re.sub(
                                    r"(测试时间[:：])[^\r\n]*",
                                    r"\g<1>" + now_str,
                                    new_text,
                                    count=1
                                )

                            try:
                                shape.TextFrame2.TextRange.Text = new_text
                            except Exception:
                                shape.TextFrame.Characters().Text = new_text

                            updated = True
                            self.log(f"文本框 9 测试时间已更新为: {now_str}")
                            break
                except Exception:
                    continue
        except Exception as e:
            self.log(f"更新文本框 9 失败: {e}")

        if not updated:
            self.log("未找到文本框 9，未更新时间")

    def delete_pictures(self, ws):
        """删除模板中的所有图片。"""
        try:
            deleted = 0
            for i in range(ws.Shapes.Count, 0, -1):
                shape = ws.Shapes(i)
                try:
                    # 11: msoLinkedPicture, 13: msoPicture
                    if shape.Type in (11, 13):
                        shape.Delete()
                        deleted += 1
                except Exception:
                    continue
            if deleted:
                self.log(f"已删除模板中原有图片 {deleted} 个")
        except Exception as e:
            self.log(f"删除原有图片失败: {e}")

    def insert_images_for_channels(self, ws, image_paths, channel_set):
        """
        将指定通道对应的截图插入到图表 8 下方，宽度维持 17cm。
        image_paths: 按顺序排列的图片路径列表，索引0对应通道0，索引1对应通道1...
        channel_set: 需要插入的通道集合（整数0-3等）
        """
        if not image_paths:
            self.log("没有可插入的截图")
            return

        try:
            chart_obj = ws.ChartObjects(8)
            left = chart_obj.Left
            top = chart_obj.Top + chart_obj.Height + 15
        except Exception as e:
            self.log(f"无法定位图表 8，截图将从 A1 附近开始插入: {e}")
            left, top = 0, 0

        width_pt = 17 * 28.35  # 17cm 转磅
        inserted_count = 0
        for ch in sorted(channel_set):
            if ch < 0 or ch >= len(image_paths):
                self.log(f"警告：通道 {ch} 超出图片数量范围，跳过")
                continue
            img_path = image_paths[ch]
            try:
                with PILImage.open(img_path) as im:
                    ratio = im.height / im.width if im.width else 1
                height_pt = width_pt * ratio

                ws.Shapes.AddPicture(
                    str(img_path),
                    False,
                    True,
                    left,
                    top,
                    width_pt,
                    height_pt
                )
                self.log(f"已插入通道 {ch} 截图: {Path(img_path).name}")
                left += width_pt + 10
                inserted_count += 1
            except Exception as e:
                self.log(f"插入通道 {ch} 截图失败: {e}")

        if inserted_count == 0:
            self.log("没有插入任何截图")

    def check_temperature_column(self, ws, last_row):
        """
        检查 temperature 列（第2行到 last_row 行）是否均为自然数，若非自然数则输出日志。
        只检查真实写入的数据行，避免把模板自带的空行/示例行算进去。
        """
        if not ws:
            return
        try:
            temp_col = None
            for c in range(1, 64):
                v = ws.Cells(1, c).Value
                if v is not None and str(v).strip().lower() == "temperature":
                    temp_col = c
                    break
            if temp_col is None:
                self.log("警告：未找到 temperature 列，跳过温度检查")
                return

            bad_rows = []
            checked = 0
            for row_idx in range(2, last_row + 1):
                val = ws.Cells(row_idx, temp_col).Value
                if val is None or str(val).strip() == "":
                    continue
                checked += 1
                num = to_number(val)
                # 自然数：整数且 >= 0
                if num is None or not num.is_integer() or num < 0:
                    coord = f"{openpyxl.utils.get_column_letter(temp_col)}{row_idx}"
                    bad_rows.append((coord, val))

            if bad_rows:
                self.log("发现 temperature 列非自然数值：")
                for coord, val in bad_rows:
                    self.log(f"  单元格 {coord}: {val}")
            else:
                self.log(f"temperature 列检查通过（{checked} 行），均为自然数")
        except Exception as e:
            self.log(f"温度列检查失败: {e}")

    # ---------------------- 数据写入 ----------------------
    @staticmethod
    def _cell_value(raw):
        """把 CSV 文本转成适合写入 Excel 的值：能转数字的转数字，空串留空，其余保持文本。"""
        if raw is None:
            return None
        text = str(raw)
        if text.strip() == "":
            return None
        stripped = text.strip()
        # 纯数字（含科学计数、负号、小数）才转数值，避免把 "2026-08-21 10:02" 这类时间拆坏
        if re.fullmatch(r"[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?", stripped):
            num = to_number(stripped)
            if num is not None:
                return num
        return text

    def write_data_rows(self, ws, data_rows, last_row):
        """一次性把数据写入工作表（单次 COM 往返，老机器上比逐单元格快很多）。"""
        if not data_rows:
            return
        n_rows = len(data_rows)
        n_cols = max(len(r) for r in data_rows)
        matrix = [[self._cell_value(r[c]) if c < len(r) else None
                   for c in range(n_cols)]
                  for r in data_rows]
        try:
            start = ws.Cells(2, 1)
            end = ws.Cells(last_row, n_cols)
            rng = ws.Range(start, end)
            rng.Value = matrix
        except Exception as e:
            self.log(f"批量写入失败，退回逐单元格写入: {e}")
            for row_idx, row_data in enumerate(data_rows, start=2):
                for col_idx, value in enumerate(row_data, start=1):
                    ws.Cells(row_idx, col_idx).Value = self._cell_value(value)

    # ---------------------- 电流校验 ----------------------
    @staticmethod
    def _series_set_current(series):
        """从图表系列名称中解析设定电流值（数字部分），失败返回 None。"""
        try:
            name = series.Name
        except Exception:
            return None
        if not name:
            return None
        m = re.search(r"(\d+(?:\.\d+)?)", str(name))
        if not m:
            return None
        return float(m.group(1))

    @staticmethod
    def _split_series_args(inner):
        """按顶层逗号切分 SERIES(...) 参数，括号与引号内的逗号不切。"""
        parts, depth, cur, in_str = [], 0, "", False
        for ch in inner:
            if ch == '"':
                in_str = not in_str
                cur += ch
            elif in_str:
                cur += ch
            elif ch == "(":
                depth += 1
                cur += ch
            elif ch == ")":
                depth -= 1
                cur += ch
            elif ch == "," and depth == 0:
                parts.append(cur.strip())
                cur = ""
            else:
                cur += ch
        if cur.strip():
            parts.append(cur.strip())
        return parts

    @classmethod
    def _series_ranges(cls, series):
        """
        解析系列 Formula，返回 (系列名, X轴区域, Y轴区域)。
        Formula 形如：=SERIES("25A",'Sheet'!$Z$2:$Z$81,'Sheet'!$X$2:$X$81,1)
        """
        try:
            formula = series.Formula
        except Exception:
            return None, None, None
        if not formula or "(" not in formula:
            return None, None, None
        f = formula.lstrip("=").strip()
        if not f.upper().startswith("SERIES"):
            return None, None, None
        inner = f[f.find("(") + 1:f.rfind(")")]
        parts = cls._split_series_args(inner)
        name = parts[0].strip('"') if parts else None
        xr = parts[1] if len(parts) > 1 else None
        yr = parts[2] if len(parts) > 2 else None
        return name, xr, yr

    @staticmethod
    def _range_first_col(range_ref):
        """
        取区域引用里的第一个列字母（如 '172K-40-BF5'!$X$2:$X$81 -> X）。
        必须先剥掉 '表名'! 前缀，否则表名里的字母（如 BF5 的 BF）会被误当成列号。
        """
        if not range_ref:
            return None
        text = str(range_ref).strip()
        # 去掉 'Sheet Name'! 或 SheetName! 前缀
        text = re.sub(r"^'[^']*'!", "", text)
        text = re.sub(r"^[A-Za-z0-9_一-龥]+!", "", text)
        m = re.match(r"\$?([A-Z]{1,3})\$?\d+", text)
        return m.group(1) if m else None

    @classmethod
    def _series_y_col_idx(cls, series):
        """取系列 Y 值（Offset 数据）所在列号。"""
        _, _, yr = cls._series_ranges(series)
        col = cls._range_first_col(yr)
        if not col:
            return None
        return cls._col_letter_to_index(col)

    @staticmethod
    def _col_letter_to_index(letter):
        n = 0
        for ch in letter:
            n = n * 26 + (ord(ch) - 64)
        return n

    @classmethod
    def _find_current_col_for_offset(cls, ws, offset_col_idx):
        """
        在表头行里找 Offset 列左侧最近的 "xx% Current" 列，即该图实际施加的负载电流列。
        参考块（第一个 Offset 列，属于 Vref 区）返回 None，不参与校验。
        """
        if offset_col_idx is None:
            return None
        if offset_col_idx == REFERENCE_BLOCK_OFFSET_COL:
            return None
        for c in range(offset_col_idx - 1, 0, -1):
            v = ws.Cells(1, c).Value
            if v is None:
                continue
            label = str(v).strip().lower()
            if label.endswith("current"):
                return c
        return None

    def _column_average(self, ws, col_idx, sample_rows):
        """取工作表指定列前 sample_rows 个非空数值的平均值。"""
        vals = []
        for row_idx in range(2, 2 + sample_rows):
            v = ws.Cells(row_idx, col_idx).Value
            if v is None or str(v).strip() == "":
                continue
            num = to_number(v)
            if num is not None:
                vals.append(num)
        if not vals:
            return None, 0
        return sum(vals) / len(vals), len(vals)

    def _compare_current(self, tag, expected, actual, col_label, rows_used):
        """比较设定电流与实际平均电流，超过容差只警告，不中断。返回是否超差。"""
        if actual is None:
            self.log(f"警告：{tag} 在列 {col_label} 前 {rows_used} 行中没有有效数值，跳过校验")
            return True
        if expected == 0:
            # 0A 无法用相对偏差衡量，只在完全非 0 时提示
            if abs(actual) > 1e-9:
                self.log(f"警告：{tag} 设定值为 0A，但列 {col_label} 前 {rows_used} 行平均值 {actual:.4f} 不为 0")
                return True
            self.log(f"{tag} 校验通过：列 {col_label} 平均值 {actual:.4f}A 与设定值 0A 一致")
            return False
        error_pct = abs(actual - expected) / abs(expected) * 100
        if error_pct > CURRENT_TOLERANCE_PCT:
            self.log(f"警告：{tag} 列 {col_label} 前 {rows_used} 行平均电流 {actual:.4f}A，"
                     f"与设定值 {expected:g}A 偏差 {error_pct:.2f}% > {CURRENT_TOLERANCE_PCT:g}%，请检查数据")
            return True
        self.log(f"{tag} 校验通过：列 {col_label} 平均 {actual:.4f}A vs 设定 {expected:g}A，偏差 {error_pct:.2f}%")
        return False

    def check_chart_series_error(self, ws, sample_rows=CURRENT_SAMPLE_ROWS):
        """
        校验图表 2~8 系列设定电流与数据前几行实际电流是否偏差在 10% 以内。

        做法：读每个系列的 SERIES(...) Formula，用 Y 值（Offset）所在列向左找最近的
        "xx% Current" 列作为该图实际施加的电流，再与该列前 sample_rows 行的平均值比较。
        同一个 (图表, Current列) 组合只校验一次，避免同一张图里重复刷屏。
        超差只输出警告日志，不抛出异常、不中断后续步骤。
        """
        warned = 0
        checked = 0

        for chart_idx in CHECK_CHART_INDEXES:
            try:
                chart = ws.ChartObjects(chart_idx).Chart
                series_count = chart.SeriesCollection().Count
            except Exception as e:
                self.log(f"警告：无法读取图表 {chart_idx}，跳过该校验（{e}）")
                continue

            # 收集本图所有 (设定电流, Current列)，按 Current 列去重
            seen = {}
            for s_idx in range(1, series_count + 1):
                try:
                    series = chart.SeriesCollection(s_idx)
                    name, _, _ = self._series_ranges(series)
                    if not name or str(name).strip().lower() == "vref":
                        continue
                    expected = self._series_set_current(series)
                    if expected is None:
                        continue

                    offset_col = self._series_y_col_idx(series)
                    current_col = self._find_current_col_for_offset(ws, offset_col)
                    if current_col is None:
                        # Vref 参考块或找不到对应 Current 列，跳过
                        continue
                    # 同一列可能被多个系列共用，保留第一个出现的设定值
                    if current_col not in seen:
                        seen[current_col] = expected
                except Exception as e:
                    self.log(f"警告：检查图表 {chart_idx} 系列 {s_idx} 时出错（{e}），继续后续校验")

            for current_col, expected in seen.items():
                col_label = openpyxl.utils.get_column_letter(current_col)
                try:
                    actual, used = self._column_average(ws, current_col, sample_rows)
                except Exception as e:
                    self.log(f"警告：读取列 {col_label} 数据失败（{e}），跳过该校验")
                    continue
                checked += 1
                over = self._compare_current(f"图表{chart_idx}", expected, actual, col_label, used)
                if over:
                    warned += 1

        if checked == 0:
            self.log("警告：未校验任何图表系列（未找到可对应的 Current 列），请检查模板结构")
        elif warned:
            self.log(f"电流校验完成：共校验 {checked} 组图表数据，其中 {warned} 组偏差超过 "
                     f"{CURRENT_TOLERANCE_PCT:g}%，已按警告处理，流程继续")
        else:
            self.log(f"电流校验完成：共校验 {checked} 组图表数据，全部在 {CURRENT_TOLERANCE_PCT:g}% 偏差范围内")

    # ---------------------- 主处理流程 ----------------------
    def process_work(self):
        excel = None
        master_wb = None
        try:
            if COM_AVAILABLE:
                try:
                    pythoncom.CoInitialize()
                except Exception:
                    pass
            else:
                self.log("需要安装 pywin32，请在命令行执行: pip install pywin32")
                return

            # 获取用户输入：文件名后缀和数据行数
            suffix = self.suffix_var.get().strip()
            if not suffix:
                suffix = "(-40_125)"
            # 清理非法字符，避免文件名错误
            suffix_safe = sanitize_filename_part(suffix)

            row_count_str = self.row_count_var.get().strip()
            try:
                row_count = int(row_count_str)
                if row_count <= 0:
                    raise ValueError
            except ValueError:
                self.log("数据行数输入无效，使用默认80行")
                row_count = 80

            # 1. 提取寄存器截图
            self.log("开始提取寄存器截图...")
            reg_images = self.extract_images_from_xlsx(self.reg_path)
            if not reg_images:
                self.log("未提取到任何图片，继续处理但不插入截图")

            # 2. 启动 Excel（整个流程只启动一次，最后统一关闭）
            self.log("正在启动 Excel...")
            try:
                excel = win32com.client.DispatchEx("Excel.Application")
            except Exception as e:
                self.log(f"启动 Excel 失败: {e}")
                return

            excel.Visible = False
            excel.DisplayAlerts = False
            # 关闭不需要的重绘/事件/自动更新，老机器上收益明显
            try:
                excel.EnableEvents = False
            except Exception:
                pass
            try:
                excel.ScreenUpdating = False
            except Exception:
                pass
            try:
                excel.AskToUpdateLinks = False
            except Exception:
                pass

            try:
                os.makedirs(self.output_dir, exist_ok=True)

                # 3. 模板只打开一次作为母版，后续每个 CSV 复制一份，避免重复加载图表
                try:
                    master_wb = excel.Workbooks.Open(self.template_path, UpdateLinks=0)
                except Exception as e:
                    self.log(f"打开模板失败: {e}")
                    return

                template_ws = master_wb.Worksheets(1)

                for csv_path in self.raw_files:
                    csv_name = Path(csv_path).name
                    stem = Path(csv_path).stem
                    safe_stem = sanitize_filename_part(stem)

                    self.log(f"\n正在处理: {csv_name}")

                    wb = None
                    try:
                        # 修改：传入 max_rows=row_count 限制读取的数据行数
                        data_rows, id_value = get_csv_data_and_id(csv_path, max_rows=row_count)
                        self.log(f"原始数据 ID = {id_value}")

                        # 计算数据最后一行行号
                        last_row = len(data_rows) + 1
                        self.log(f"数据共 {len(data_rows)} 行，Excel 最后一行行号为 {last_row}")

                        # 提取 B 列（索引1）的通道值（去重且为整数0-3）
                        channel_set = set()
                        for row in data_rows:
                            if len(row) > 1 and row[1] and str(row[1]).strip():
                                try:
                                    ch = int(float(str(row[1]).strip()))
                                    if 0 <= ch <= 3:
                                        channel_set.add(ch)
                                except ValueError:
                                    pass
                        self.log(f"数据中出现的通道: {sorted(channel_set)}")

                        # 用母版复制一份工作簿副本，比重新 Open 模板快得多（不重复解析图表）
                        template_ws.Copy()
                        wb = excel.ActiveWorkbook
                        ws = wb.Worksheets(1)

                        # 3.1 清空并写入所有数据行
                        # 至少清除到第81行（默认80行数据对应的下一行），若数据行数更多则清除到实际最后一行之后
                        clear_end = max(81, last_row)
                        try:
                            ws.Rows(f"2:{clear_end}").ClearContents()
                        except Exception as e:
                            self.log(f"清除旧数据失败: {e}")

                        # 一次性写入（单次 COM 往返）
                        self.write_data_rows(ws, data_rows, last_row)

                        # 3.2 修改图表 8 标题
                        self.set_chart8_title(ws, id_value)

                        # 3.3 更新时间文本框 9
                        self.update_textbox9_time(ws)

                        # 3.4 校验图表2~8 系列设定电流与数据前几行实际电流（偏差>10% 仅警告）
                        self.check_chart_series_error(ws, sample_rows=CURRENT_SAMPLE_ROWS)

                        # 3.5 检查 temperature 列（真实数据行范围内）
                        self.check_temperature_column(ws, last_row)

                        # 3.6 删除原有图片
                        self.delete_pictures(ws)

                        # 3.7 插入匹配通道的截图
                        if reg_images:
                            self.insert_images_for_channels(ws, reg_images, channel_set)
                        else:
                            self.log("没有图片可插入")

                        # 3.8 保存到临时文件
                        temp_out_path = os.path.join(
                            self.output_dir,
                            f"{safe_stem}_temp_{datetime.now().strftime('%Y%m%d%H%M%S%f')}.xlsx"
                        )
                        self.log(f"临时保存到: {temp_out_path}")
                        wb.SaveAs(temp_out_path, FileFormat=51)  # 51 = xlsx
                        wb.Close(SaveChanges=False)
                        wb = None

                        # 3.9 读取新文件中的 ID 并重命名
                        final_id = get_id_from_excel_file(temp_out_path)
                        if final_id is None:
                            final_id = id_value
                            self.log(f"从新文件读取 ID 失败，使用 CSV 中的 ID: {final_id}")

                        # 修改：使用用户输入的后缀
                        if final_id:
                            safe_id = sanitize_filename_part(final_id)
                            final_out_path = os.path.join(self.output_dir, f"{safe_stem}_{safe_id}_{suffix_safe}.xlsx")
                        else:
                            final_out_path = os.path.join(self.output_dir, f"{safe_stem}_{suffix_safe}.xlsx")

                        # 若最终文件已存在则删除
                        if os.path.exists(final_out_path):
                            try:
                                os.remove(final_out_path)
                                self.log(f"已删除旧文件: {final_out_path}")
                            except PermissionError:
                                self.log(f"无法删除被占用的文件: {final_out_path}")
                                try:
                                    os.remove(temp_out_path)
                                except Exception:
                                    pass
                                continue

                        try:
                            os.rename(temp_out_path, final_out_path)
                            self.log(f"已生成: {final_out_path}")
                        except Exception as e:
                            self.log(f"重命名文件失败: {e}")
                            try:
                                os.remove(temp_out_path)
                            except Exception:
                                pass

                    except Exception as e:
                        self.log(f"处理 {csv_name} 时出错: {e}")
                    finally:
                        # 确保当前工作簿已关闭（母版保持打开状态供下一个文件复用）
                        if wb is not None:
                            try:
                                wb.Close(SaveChanges=False)
                            except Exception:
                                pass
                            wb = None

                # 4. 全部处理完成后再关闭母版与 Excel
                if master_wb is not None:
                    try:
                        master_wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                    master_wb = None

            finally:
                if excel is not None:
                    try:
                        excel.Quit()
                    except Exception:
                        pass
                    excel = None

        finally:
            # 清理临时截图
            if self.temp_dir and os.path.isdir(self.temp_dir):
                shutil.rmtree(self.temp_dir, ignore_errors=True)
                self.temp_dir = None
            if COM_AVAILABLE:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

        self.root.after(0, lambda: messagebox.showinfo("完成", "全部处理完成"))
        self.log("全部完成")


if __name__ == "__main__":
    app = App()
    app.root.mainloop()
