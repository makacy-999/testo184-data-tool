#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Testo 184 温度计数据读取汇总工具 - 后端服务
Flask + SQLite + openpyxl
"""

import os
import sys
import csv
import json
import uuid
import platform
import subprocess
import sqlite3
import threading
from datetime import datetime
from pathlib import Path
from flask import Flask, request, jsonify, render_template, send_file

# .vi2 专有格式解析（Testo ComSoft 存档）
try:
    from vi2_parser import parse_vi2
    VI2_AVAILABLE = True
except ImportError:
    VI2_AVAILABLE = False

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("❌ 缺少 openpyxl，请运行: pip install openpyxl")
    sys.exit(1)

# PDF 报告解析（Testo 184 内置 measurement report.pdf）
try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False


# ─── 配置 ────────────────────────────────────────────────────────────────────

# 兼容 PyInstaller 打包：frozen 模式下资源解压到 sys._MEIPASS
if getattr(sys, "frozen", False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 数据库和导出目录放在用户主目录，打包后也可写
APP_DATA_DIR = os.path.join(os.path.expanduser("~"), ".testo184_data")
DB_PATH = os.path.join(APP_DATA_DIR, "testo_data.db")
EXPORT_DIR = os.path.join(APP_DATA_DIR, "exports")
os.makedirs(EXPORT_DIR, exist_ok=True)


# ─── Flask 初始化 ────────────────────────────────────────────────────────────

app = Flask(
    __name__,
    static_folder=os.path.join(BASE_DIR, "static"),
    template_folder=os.path.join(BASE_DIR, "templates"),
)


# ─── 数据库 ──────────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS sessions (
            id TEXT PRIMARY KEY,
            name TEXT,
            created_at TEXT,
            total_points INTEGER DEFAULT 0,
            completed_points INTEGER DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS measurement_points (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id TEXT NOT NULL,
            point_number TEXT NOT NULL,
            serial_number TEXT DEFAULT '',
            sort_order INTEGER,
            FOREIGN KEY (session_id) REFERENCES sessions(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS device_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id TEXT NOT NULL,
            point_number TEXT,
            serial_number TEXT,
            device_name TEXT,
            csv_file TEXT,
            record_index INTEGER,
            date_val TEXT,
            time_val TEXT,
            temperature REAL,
            humidity REAL,
            alarm TEXT,
            raw_data TEXT,
            created_at TEXT,
            FOREIGN KEY (session_id) REFERENCES sessions(id) ON DELETE CASCADE
        );
    """)
    conn.commit()
    conn.close()


# 内存中跟踪当前会话状态
session_lock = threading.Lock()
running_sessions = {}


# ─── 设备检测 ────────────────────────────────────────────────────────────────

class DeviceDetector:
    """跨平台 USB 设备检测"""

    @staticmethod
    def get_mount_points():
        system = platform.system()
        points = []
        if system == "Darwin":
            volumes = Path("/Volumes")
            if volumes.exists():
                for item in volumes.iterdir():
                    if item.is_dir() and not item.name.startswith("."):
                        points.append(str(item))
        elif system == "Windows":
            try:
                import string
                for letter in string.ascii_uppercase:
                    drive = f"{letter}:\\"
                    if os.path.exists(drive):
                        points.append(drive)
            except Exception:
                try:
                    result = subprocess.run(
                        ["powershell", "-Command",
                         "Get-PSDrive -PSProvider FileSystem | Select-Object -ExpandProperty Root"],
                        capture_output=True, text=True, timeout=10
                    )
                    for line in result.stdout.strip().split("\n"):
                        line = line.strip()
                        if line:
                            points.append(line)
                except Exception:
                    pass
        else:  # Linux
            for base in ["/media", "/mnt"]:
                base_path = Path(base)
                if base_path.exists():
                    for item in base_path.rglob("*"):
                        if item.is_dir() and item.parent != base_path or item.is_dir():
                            points.append(str(item))
        return points

    @staticmethod
    def is_testo_device(path):
        """识别是否为 Testo 184 温度计设备。
        放宽判断：不要求挂载名必须含 TESTO/184（很多情况设备名是随机的），
        改为优先读取目录下 CSV 内容特征来判断。"""
        name = os.path.basename(path).upper()
        if "TESTO" in name or "184" in name:
            return True
        try:
            files = os.listdir(path)
        except (OSError, PermissionError):
            return False

        # 读前几个 CSV 的内容特征
        csv_files = [f for f in files if f.lower().endswith(".csv")]
        for f in csv_files[:5]:
            fp = os.path.join(path, f)
            try:
                with open(fp, "rb") as fh:
                    head = fh.read(3000).decode("utf-8", "ignore").lower()
                if "testo" in head:
                    return True
            except Exception:
                continue

        files_lower = [f.lower() for f in files]
        testo_keywords = ["testo", "configuration", "measurement", "messdaten", "184"]
        return sum(1 for kw in testo_keywords if any(kw in f for f in files_lower)) >= 2

    @classmethod
    def scan(cls):
        devices = []
        for mp in cls.get_mount_points():
            if cls.is_testo_device(mp):
                device = {
                    "path": mp,
                    "name": os.path.basename(mp),
                    "csv_files": [],
                    "records": [],
                    "serial_number": "",
                    "error": None,
                }
                csv_files = []
                pdf_files = []
                vi2_files = []
                try:
                    for root, dirs, files in os.walk(mp):
                        # 不排除隐藏目录：Testo 设备数据文件可能位于隐藏/系统目录（如 .SystemVolumeInformation）
                        dirs[:] = [d for d in dirs
                                   if d.lower() not in ("found.000", "system volume information")
                                   and d.lower() != "$recycle.bin"]
                        for f in files:
                            fl = f.lower()
                            if fl.endswith(".csv"):
                                csv_files.append(os.path.join(root, f))
                            elif fl.endswith(".pdf") or "report" in fl or "measurement" in fl:
                                pdf_files.append(os.path.join(root, f))
                            elif fl.endswith(".vi2"):
                                vi2_files.append(os.path.join(root, f))
                            else:
                                # 无扩展名的报告文件：按内容签名识别（Testo 设备报告常无扩展名）
                                fp = os.path.join(root, f)
                                try:
                                    with open(fp, "rb") as fh:
                                        sig = fh.read(2048)
                                    if sig.startswith(b"%PDF"):
                                        pdf_files.append(fp)
                                    elif sig.startswith(b"D0CF11E0"):
                                        vi2_files.append(fp)
                                except Exception:
                                    pass
                except (OSError, PermissionError) as e:
                    device["error"] = f"读取目录失败: {e}"

                device["csv_files"] = csv_files
                device["pdf_files"] = [os.path.basename(f) for f in pdf_files]
                device["vi2_files"] = [os.path.basename(f) for f in vi2_files]
                all_records = []
                device_info = {}

                for csv_file in csv_files:
                    try:
                        headers, records, info = cls.parse_csv(csv_file)
                        for r in records:
                            r["source_file"] = os.path.basename(csv_file)
                        all_records.extend(records)
                        device_info.update(info)
                    except Exception as e:
                        device["error"] = f"解析 {os.path.basename(csv_file)} 失败: {e}"

                # 如果没有 CSV，或 CSV 为空，尝试从 Testo 报告 PDF 中提取数据（内置报告）
                if not all_records:
                    for pdf_file in pdf_files:
                        try:
                            headers, records, info = cls.parse_testo_pdf(pdf_file)
                            if records:
                                for r in records:
                                    r["source_file"] = os.path.basename(pdf_file)
                                all_records.extend(records)
                                device_info.update(info)
                                break
                        except Exception as e:
                            device["error"] = f"解析 {os.path.basename(pdf_file)} 失败: {e}"

                # 若仍无数据，尝试解析 .vi2 专有存档（Testo ComSoft 导出格式）
                if not all_records and VI2_AVAILABLE:
                    for vi2_file in vi2_files:
                        try:
                            parsed = parse_vi2(vi2_file)
                            if parsed.get("records"):
                                device_info["serial"] = parsed["serial_number"]
                                device_info["unit"] = parsed["unit"]
                                for idx, r in enumerate(parsed["records"], 1):
                                    all_records.append({
                                        "date": r["date"],
                                        "time": r["time"],
                                        "temperature": r["temperature"],
                                        "humidity": None,
                                        "alarm": "",
                                        "_raw": {"t_code": r["t_code"]},
                                        "source_file": os.path.basename(vi2_file),
                                        "record_index": idx,
                                    })
                                device_info["_vi2_sample_minutes"] = parsed["sample_minutes"]
                                break
                        except Exception as e:
                            device["error"] = f"解析 {os.path.basename(vi2_file)} 失败: {e}"

                device["records"] = all_records
                # 尝试提取序列号：从 CSV 注释行 / 文件名 / 首行元信息
                for key in ["serial", "serial_number", "sn", "s/n", "序列号", "deviceserialnumber", "serial no."]:
                    if key in device_info and device_info[key]:
                        device["serial_number"] = str(device_info[key])
                        break
                if not device["serial_number"]:
                    # 从 CSV 文件名中提取（Testo 184 文件名常含 SN，如 184XXXX.csv）
                    import re as _re
                    for cf in device.get("csv_files", []):
                        base = os.path.basename(cf)
                        m = _re.search(r"(\d{6,})", base)
                        if m:
                            device["serial_number"] = m.group(1)
                            break
                # 仍然没找到：扫描设备目录下所有非 CSV 的元信息文件首行
                if not device["serial_number"]:
                    try:
                        for root, _, fs in os.walk(mp):
                            for f in fs:
                                if f.lower().endswith((".txt", ".ini", ".log", ".cfg")):
                                    fp = os.path.join(root, f)
                                    for enc in ["utf-8-sig", "utf-8", "latin-1", "gbk"]:
                                        try:
                                            with open(fp, "r", encoding=enc) as fh:
                                                first = fh.read(2000)
                                            import re as _re
                                            m = _re.search(r"(?:serial|sn|nr)\.?:?\s*([A-Za-z0-9\-]{4,20})",
                                                            first, _re.IGNORECASE)
                                            if m:
                                                device["serial_number"] = m.group(1).strip()
                                            break
                                        except (UnicodeDecodeError, OSError):
                                            continue
                                    if device["serial_number"]:
                                        break
                            if device["serial_number"]:
                                break
                    except Exception:
                        pass
                # 如果实在提取不到，拼接设备名作为标识（用户可在界面手动修正）
                if not device["serial_number"]:
                    device["serial_number"] = device["name"]

                devices.append(device)
        return devices

    @staticmethod
    def parse_csv(csv_path):
        """解析 CSV 文件，返回 (headers, records, device_info)"""
        records = []
        device_info = {}
        encodings = ["utf-8-sig", "utf-8", "latin-1", "cp1252", "gbk"]
        content = None

        for enc in encodings:
            try:
                with open(csv_path, "r", encoding=enc) as f:
                    content = f.read()
                break
            except (UnicodeDecodeError, UnicodeError):
                continue

        if content is None:
            raise ValueError(f"无法读取文件编码: {csv_path}")

        lines = content.splitlines()
        if not lines:
            raise ValueError("CSV 文件为空")

        # 跳过注释行，提取设备信息
        data_start = 0
        for i, line in enumerate(lines):
            stripped = line.strip()
            if stripped.startswith("#") or stripped.startswith("//"):
                if ":" in stripped:
                    key, _, val = stripped.lstrip("#/ ").partition(":")
                    device_info[key.strip().lower()] = val.strip()
                data_start = i + 1
            else:
                break

        data_lines = lines[data_start:]
        if not data_lines:
            raise ValueError("无数据行")

        # 自动检测分隔符
        sample = "\n".join(data_lines[:10])
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            sep = dialect.delimiter
        except csv.Error:
            sep = ","

        reader = csv.reader(data_lines, delimiter=sep)
        rows = list(reader)
        if not rows:
            raise ValueError("CSV 解析后无数据")

        # 查找表头行
        header_idx = 0
        keywords = ["temp", "温度", "temperature", "°c", "°f",
                    "date", "日期", "time", "时间", "humidity", "湿度"]
        for i, row in enumerate(rows[:10]):
            text = " ".join(c.lower() for c in row)
            if any(kw in text for kw in keywords):
                header_idx = i
                break

        headers = [h.strip() for h in rows[header_idx]]
        data_rows = rows[header_idx + 1:]

        # 列映射
        col_map = {}
        for idx, h in enumerate(headers):
            hl = h.lower()
            if any(k in hl for k in ["date", "日期", "datum"]):
                col_map["date"] = idx
            elif any(k in hl for k in ["time", "时间", "zeit"]):
                col_map["time"] = idx
            elif any(k in hl for k in ["temp", "温度", "temperature"]):
                col_map["temperature"] = idx
            elif any(k in hl for k in ["humid", "湿度", "feuchte"]):
                col_map["humidity"] = idx
            elif any(k in hl for k in ["alarm", "报警"]):
                col_map["alarm"] = idx

        for row in data_rows:
            if not row or all(not c.strip() for c in row):
                continue
            rec = {}
            for field, ci in col_map.items():
                if ci < len(row):
                    rec[field] = row[ci].strip()
            # 原始数据
            rec["_raw"] = {headers[i]: (row[i].strip() if i < len(row) else "")
                           for i in range(len(headers))}
            if rec:
                records.append(rec)

        device_info["csv_file"] = os.path.basename(csv_path)
        device_info["csv_path"] = csv_path
        return headers, records, device_info

    @staticmethod
    def parse_testo_pdf(pdf_path):
        """解析 Testo 184 内置的 measurement report PDF，提取温度数据。
        Testo 报告 PDF 内通常含测量数据表格（日期/时间/温度/湿度/报警）"""
        records = []
        device_info = {}
        if not PDFPLUMBER_AVAILABLE:
            raise RuntimeError("未安装 pdfplumber，无法解析 PDF 报告")

        headers_out = None
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:10]:  # 最多看前10页
                tables = page.extract_tables()
                for table in tables:
                    if not table or not table[0]:
                        continue
                    # 尝试定位表头行
                    header_idx = 0
                    for i, row in enumerate(table[:4]):
                        joined = " ".join(str(c).lower() for c in row if c)
                        if any(k in joined for k in ["date", "日期", "datum"]):
                            header_idx = i
                            break
                    headers = [str(c).strip() if c else "" for c in table[header_idx]]
                    # 列映射
                    col_map = {}
                    for idx, h in enumerate(headers):
                        hl = h.lower()
                        if "date" in hl or "日期" in hl:
                            col_map["date"] = idx
                        elif "time" in hl or "时间" in hl:
                            col_map["time"] = idx
                        elif "temp" in hl or "温度" in hl:
                            col_map["temperature"] = idx
                        elif "humid" in hl or "湿度" in hl:
                            col_map["humidity"] = idx
                        elif "alarm" in hl or "报警" in hl or "grenzwert" in hl:
                            col_map["alarm"] = idx
                    if not col_map.get("temperature") and not col_map.get("humidity"):
                        continue
                    headers_out = headers
                    for row in table[header_idx + 1:]:
                        if not row or all(not c for c in row):
                            continue
                        rec = {}
                        for field, ci in col_map.items():
                            if ci < len(row):
                                rec[field] = str(row[ci]).strip() if row[ci] is not None else ""
                        if rec and (rec.get("temperature") or rec.get("humidity")):
                            rec["_raw"] = {headers[i]: (str(row[i]).strip() if i < len(row) and row[i] is not None else "")
                                           for i in range(len(headers))}
                            records.append(rec)

        if not records:
            # 回退：表格提取失败时，用文本行解析（针对无表格线的 PDF）
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    for page in pdf.pages[:10]:
                        text = page.extract_text() or ""
                        for line in text.splitlines():
                            line_s = line.strip()
                            if not line_s:
                                continue
                            import re
                            # 例: 2026-09-20 08:00:00 5.2 42.0 Normal
                            m = re.match(
                                r"^(\d{4}[-/]\d{1,2}[-/]\d{1,2})"
                                r"\s+(\d{1,2}:\d{2}(?::\d{2})?)"
                                r"\s+([-\d.]+\s*°?C?)\s+([-\d.]+)"
                                r"(?:\s+(.+))?$", line_s)
                            if m:
                                records.append({
                                    "date": m.group(1),
                                    "time": m.group(2),
                                    "temperature": m.group(3).replace("°C", "").strip(),
                                    "humidity": m.group(4),
                                    "alarm": m.group(5) or "",
                                    "_raw": {"Temperature": m.group(3), "Humidity": m.group(4)},
                                })
            except Exception:
                pass

        if not records:
            raise ValueError("PDF 中未找到测量数据表格")

        device_info["pdf_file"] = os.path.basename(pdf_path)
        device_info["source"] = "pdf_report"
        return headers_out or [], records, device_info


# ─── Excel 导出 ─────────────────────────────────────────────────────────────

def export_excel(session_id, selected_indexes=None):
    """将指定会话的数据导出为 Excel（selected_indexes 为可选勾选的测点列表）"""
    conn = get_db()
    session = conn.execute("SELECT * FROM sessions WHERE id=?", (session_id,)).fetchone()
    if not session:
        conn.close()
        return None, "会话不存在"

    points = conn.execute(
        "SELECT * FROM measurement_points WHERE session_id=? ORDER BY sort_order",
        (session_id,)
    ).fetchall()

    # 若指定了勾选的测点下标，则仅导出这些测点
    if selected_indexes:
        points = [points[int(i)] for i in selected_indexes if int(i) < len(points)]

    HEADER_FONT = Font(bold=True, size=11, color="FFFFFF")
    HEADER_FILL = PatternFill(start_color="2F75B5", end_color="2F75B5", fill_type="solid")
    HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
    DATA_ALIGN = Alignment(horizontal="center", vertical="center")
    BORDER = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    wb = openpyxl.Workbook()

    # ── Sheet 1: 汇总数据 ──
    ws = wb.active
    ws.title = "汇总数据"
    base_headers = ["测点号", "设备SN号", "序号", "日期", "时间", "温度(°C)", "湿度(%)", "报警"]

    for col, h in enumerate(base_headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = BORDER

    row_num = 2
    for pt in points:
        records = conn.execute(
            "SELECT * FROM device_records WHERE session_id=? AND point_number=? ORDER BY record_index",
            (session_id, pt["point_number"])
        ).fetchall()
        for rec in records:
            ws.cell(row=row_num, column=1, value=pt["point_number"])
            ws.cell(row=row_num, column=2, value=pt["serial_number"] or rec["serial_number"] or "")
            ws.cell(row=row_num, column=3, value=rec["record_index"])
            ws.cell(row=row_num, column=4, value=rec["date_val"] or "")
            ws.cell(row=row_num, column=5, value=rec["time_val"] or "")
            ws.cell(row=row_num, column=6, value=rec["temperature"])
            ws.cell(row=row_num, column=7, value=rec["humidity"])
            ws.cell(row=row_num, column=8, value=rec["alarm"] or "")
            for c in range(1, len(base_headers) + 1):
                cell = ws.cell(row=row_num, column=c)
                cell.alignment = DATA_ALIGN
                cell.border = BORDER
            row_num += 1

    for i, w in enumerate([12, 18, 8, 14, 12, 12, 10, 12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if row_num > 2:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(base_headers))}{row_num - 1}"

    # ── Sheet 2: 设备概览 ──
    ws2 = wb.create_sheet("设备概览")
    s2_headers = ["测点号", "设备SN号", "数据点数", "最早时间", "最晚时间"]
    for col, h in enumerate(s2_headers, 1):
        cell = ws2.cell(row=1, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = BORDER

    for row_idx, pt in enumerate(points, 2):
        records = conn.execute(
            "SELECT * FROM device_records WHERE session_id=? AND point_number=? ORDER BY date_val, time_val",
            (session_id, pt["point_number"])
        ).fetchall()
        dates = [r["date_val"] for r in records if r["date_val"]]
        ws2.cell(row=row_idx, column=1, value=pt["point_number"])
        ws2.cell(row=row_idx, column=2, value=pt["serial_number"] or "")
        ws2.cell(row=row_idx, column=3, value=len(records))
        ws2.cell(row=row_idx, column=4, value=min(dates) if dates else "")
        ws2.cell(row=row_idx, column=5, value=max(dates) if dates else "")
        for c in range(1, len(s2_headers) + 1):
            ws2.cell(row=row_idx, column=c).alignment = DATA_ALIGN
            ws2.cell(row=row_idx, column=c).border = BORDER

    for i, w in enumerate([12, 18, 12, 14, 14], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = "A2"

    # ── Sheet 3: 统计摘要 ──
    ws3 = wb.create_sheet("统计摘要")
    s3_headers = ["测点号", "设备SN号", "最小温度(°C)", "最大温度(°C)", "平均温度(°C)", "数据点数"]
    for col, h in enumerate(s3_headers, 1):
        cell = ws3.cell(row=1, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = BORDER

    for row_idx, pt in enumerate(points, 2):
        records = conn.execute(
            "SELECT temperature FROM device_records WHERE session_id=? AND point_number=? AND temperature IS NOT NULL",
            (session_id, pt["point_number"])
        ).fetchall()
        temps = [r["temperature"] for r in records if r["temperature"] is not None]
        ws3.cell(row=row_idx, column=1, value=pt["point_number"])
        ws3.cell(row=row_idx, column=2, value=pt["serial_number"] or "")
        ws3.cell(row=row_idx, column=3, value=round(min(temps), 1) if temps else "")
        ws3.cell(row=row_idx, column=4, value=round(max(temps), 1) if temps else "")
        ws3.cell(row=row_idx, column=5, value=round(sum(temps) / len(temps), 1) if temps else "")
        ws3.cell(row=row_idx, column=6, value=len(records))
        for c in range(1, len(s3_headers) + 1):
            ws3.cell(row=row_idx, column=c).alignment = DATA_ALIGN
            ws3.cell(row=row_idx, column=c).border = BORDER

    for i, w in enumerate([12, 18, 15, 15, 15, 12], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    ws3.freeze_panes = "A2"

    # ── Sheet 4+：每台设备单独一张完整数据表（按分钟记录的时间序列）──
    used_names = {"汇总数据", "设备概览", "统计摘要"}
    for pt in points:
        records = conn.execute(
            "SELECT * FROM device_records WHERE session_id=? AND point_number=? "
            "ORDER BY date_val, time_val, record_index",
            (session_id, pt["point_number"])
        ).fetchall()
        if not records:
            continue

        # Sheet 名：测点{N}，处理重名/超长/非法字符
        base_name = f"测点{pt['point_number']}"
        sheet_name = base_name
        n = 2
        while sheet_name in used_names:
            sheet_name = f"{base_name}-{n}"
            n += 1
        sheet_name = sheet_name[:31]
        used_names.add(sheet_name)

        wsd = wb.create_sheet(sheet_name)
        s_headers = ["序号", "日期", "时间", "温度(°C)", "湿度(%)", "报警"]
        for col, h in enumerate(s_headers, 1):
            cell = wsd.cell(row=1, column=col, value=h)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = HEADER_ALIGN
            cell.border = BORDER

        for row_idx, rec in enumerate(records, 2):
            wsd.cell(row=row_idx, column=1, value=rec["record_index"])
            wsd.cell(row=row_idx, column=2, value=rec["date_val"] or "")
            wsd.cell(row=row_idx, column=3, value=rec["time_val"] or "")
            wsd.cell(row=row_idx, column=4, value=rec["temperature"])
            wsd.cell(row=row_idx, column=5, value=rec["humidity"])
            wsd.cell(row=row_idx, column=6, value=rec["alarm"] or "")
            for c in range(1, len(s_headers) + 1):
                wsd.cell(row=row_idx, column=c).alignment = DATA_ALIGN
                wsd.cell(row=row_idx, column=c).border = BORDER

        for i, w in enumerate([8, 14, 12, 12, 10, 12], 1):
            wsd.column_dimensions[get_column_letter(i)].width = w
        wsd.freeze_panes = "A2"
        if len(records) > 0:
            wsd.auto_filter.ref = f"A1:{get_column_letter(len(s_headers))}{len(records) + 1}"

    # 保存
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Testo184_汇总_{ts}.xlsx"
    filepath = os.path.join(EXPORT_DIR, filename)
    wb.save(filepath)
    conn.close()
    return filepath, filename


# ─── 前端页面 ─────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


# ─── API: 会话管理 ────────────────────────────────────────────────────────────

@app.route("/api/sessions", methods=["POST"])
def create_session():
    data = request.json
    name = data.get("name", "") or f"测试会话 {datetime.now().strftime('%m-%d %H:%M')}"
    points = data.get("points", [])  # [{"point_number": "1"}, ...]

    session_id = str(uuid.uuid4())[:8]
    conn = get_db()
    conn.execute(
        "INSERT INTO sessions (id, name, total_points, completed_points) VALUES (?, ?, ?, 0)",
        (session_id, name, len(points))
    )
    for i, pt in enumerate(points):
        conn.execute(
            "INSERT INTO measurement_points (session_id, point_number, sort_order) VALUES (?, ?, ?)",
            (session_id, str(pt["point_number"]), i)
        )
    conn.commit()
    conn.close()

    with session_lock:
        running_sessions[session_id] = {
            "current_point_index": 0,
            "detected_path": None,
        }

    return jsonify({
        "id": session_id, "name": name,
        "total_points": len(points), "completed_points": 0,
        "points": [{"point_number": str(pt["point_number"]), "serial_number": "", "status": "pending"}
                   for pt in points]
    })


@app.route("/api/sessions", methods=["GET"])
def list_sessions():
    conn = get_db()
    sessions = conn.execute("SELECT * FROM sessions ORDER BY created_at DESC LIMIT 50").fetchall()
    result = []
    for s in sessions:
        pts = conn.execute(
            "SELECT * FROM measurement_points WHERE session_id=? ORDER BY sort_order",
            (s["id"],)
        ).fetchall()
        total_records = conn.execute(
            "SELECT COUNT(*) as cnt FROM device_records WHERE session_id=?", (s["id"],)
        ).fetchone()["cnt"]
        result.append({
            "id": s["id"], "name": s["name"],
            "total_points": s["total_points"],
            "completed_points": s["completed_points"],
            "created_at": s["created_at"],
            "total_records": total_records,
            "points": [dict(p) for p in pts],
        })
    conn.close()
    return jsonify(result)


@app.route("/api/sessions/<session_id>", methods=["GET"])
def get_session(session_id):
    conn = get_db()
    s = conn.execute("SELECT * FROM sessions WHERE id=?", (session_id,)).fetchone()
    if not s:
        conn.close()
        return jsonify({"error": "会话不存在"}), 404

    pts = conn.execute(
        "SELECT * FROM measurement_points WHERE session_id=? ORDER BY sort_order",
        (session_id,)
    ).fetchall()

    points_data = []
    for p in pts:
        rec_count = conn.execute(
            "SELECT COUNT(*) as cnt FROM device_records WHERE session_id=? AND point_number=?",
            (session_id, p["point_number"])
        ).fetchone()["cnt"]
        points_data.append({
            **dict(p),
            "record_count": rec_count,
            "status": "completed" if rec_count > 0 else "pending"
        })

    total_records = conn.execute(
        "SELECT COUNT(*) as cnt FROM device_records WHERE session_id=?", (session_id,)
    ).fetchone()["cnt"]
    conn.close()

    return jsonify({
        "id": s["id"], "name": s["name"],
        "total_points": s["total_points"],
        "completed_points": s["completed_points"],
        "total_records": total_records,
        "points": points_data,
    })


@app.route("/api/sessions/<session_id>", methods=["DELETE"])
def delete_session(session_id):
    conn = get_db()
    conn.execute("DELETE FROM sessions WHERE id=?", (session_id,))
    conn.commit()
    conn.close()
    with session_lock:
        running_sessions.pop(session_id, None)
    return jsonify({"ok": True})


# ─── API: 测点管理 ────────────────────────────────────────────────────────────

@app.route("/api/sessions/<session_id>/points", methods=["PUT"])
def update_points(session_id):
    data = request.json
    points = data.get("points", [])

    conn = get_db()
    conn.execute("DELETE FROM measurement_points WHERE session_id=?", (session_id,))
    for i, pt in enumerate(points):
        conn.execute(
            "INSERT INTO measurement_points (session_id, point_number, serial_number, sort_order) "
            "VALUES (?, ?, ?, ?)",
            (session_id, str(pt["point_number"]), pt.get("serial_number", ""), i)
        )
    conn.execute("UPDATE sessions SET total_points=? WHERE id=?", (len(points), session_id))
    conn.commit()
    conn.close()

    return jsonify({"ok": True})


# ─── API: 设备检测与读取 ──────────────────────────────────────────────────────

@app.route("/api/sessions/<session_id>/detect", methods=["POST"])
def detect_device(session_id):
    """检测当前插入的 Testo 设备"""
    try:
        devices = DeviceDetector.scan()
    except Exception as e:
        return jsonify({"error": f"检测失败: {e}"}), 500

    if not devices:
        return jsonify({"detected": False, "message": "未检测到 Testo 184 设备，请确认设备已通过 USB 连接"})

    # 检查哪些设备还没读过
    conn = get_db()
    points = conn.execute(
        "SELECT * FROM measurement_points WHERE session_id=? ORDER BY sort_order",
        (session_id,)
    ).fetchall()
    read_sns = set()
    for p in points:
        if p["serial_number"]:
            read_sns.add(p["serial_number"])
    conn.close()

    # 找到未读的设备
    for dev in devices:
        sn = dev.get("serial_number", dev["name"])
        if sn not in read_sns:
            # 获取数据预览（前20条）
            preview = dev["records"][:20]
            return jsonify({
                "detected": True,
                "device": {
                    "name": dev["name"],
                    "path": dev["path"],
                    "serial_number": sn,
                    "record_count": len(dev["records"]),
                    "csv_files": [os.path.basename(f) for f in dev["csv_files"]],
                    "pdf_files": dev.get("pdf_files", []),
                    "source": "CSV" if dev["csv_files"] else ("PDF报告" if dev.get("pdf_files") else "无"),
                    "preview": preview,
                    "error": dev.get("error"),
                }
            })

    return jsonify({
        "detected": False,
        "message": "所有已检测到的设备都已读取过，请插入新的设备"
    })


@app.route("/api/sessions/<session_id>/read", methods=["POST"])
def read_device(session_id):
    """读取当前检测到的设备并保存数据"""
    data = request.json or {}
    device_path = data.get("device_path")
    serial_number = data.get("serial_number", "")

    if not device_path:
        return jsonify({"error": "缺少设备路径"}), 400

    # 找到当前待读取的测点
    conn = get_db()
    next_point = conn.execute(
        "SELECT * FROM measurement_points WHERE session_id=? AND (serial_number='' OR serial_number IS NULL) "
        "ORDER BY sort_order LIMIT 1",
        (session_id,)
    ).fetchone()

    if not next_point:
        conn.close()
        return jsonify({"error": "所有测点都已关联设备"}), 400

    # 扫描并读取设备
    try:
        devices = DeviceDetector.scan()
    except Exception as e:
        conn.close()
        return jsonify({"error": f"读取设备失败: {e}"}), 500

    target_device = None
    for dev in devices:
        if dev["path"] == device_path:
            target_device = dev
            break

    if not target_device:
        conn.close()
        return jsonify({"error": "设备已断开连接"}), 400

    sn = serial_number or target_device.get("serial_number", target_device["name"])
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 保存记录
    for idx, rec in enumerate(target_device["records"], 1):
        temp = None
        try:
            t = rec.get("temperature", "").replace(",", ".").replace("°C", "").replace("°F", "").strip()
            temp = float(t)
        except (ValueError, AttributeError):
            pass

        hum = None
        try:
            h = rec.get("humidity", "").replace(",", ".").replace("%", "").strip()
            hum = float(h)
        except (ValueError, AttributeError):
            pass

        raw = rec.get("_raw", {})
        conn.execute(
            "INSERT INTO device_records "
            "(session_id, point_number, serial_number, device_name, csv_file, "
            " record_index, date_val, time_val, temperature, humidity, alarm, raw_data, created_at) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (session_id, next_point["point_number"], sn, target_device["name"],
             rec.get("source_file", ""), idx,
             rec.get("date", ""), rec.get("time", ""), temp, hum,
             rec.get("alarm", ""), json.dumps(raw, ensure_ascii=False), now)
        )

    # 更新测点的 SN
    conn.execute(
        "UPDATE measurement_points SET serial_number=? WHERE id=?",
        (sn, next_point["id"])
    )
    # 更新已完成数
    completed = conn.execute(
        "SELECT COUNT(*) as cnt FROM measurement_points WHERE session_id=? AND serial_number!='' AND serial_number IS NOT NULL",
        (session_id,)
    ).fetchone()["cnt"]
    conn.execute("UPDATE sessions SET completed_points=? WHERE id=?", (completed, session_id))
    conn.commit()

    record_count = len(target_device["records"])
    conn.close()

    return jsonify({
        "ok": True,
        "point_number": next_point["point_number"],
        "serial_number": sn,
        "record_count": record_count,
        "completed_points": completed,
        "message": f"测点 {next_point['point_number']} 读取完成，共 {record_count} 条数据"
    })


# ─── API: 导入 .vi2 文件 ─────────────────────────────────────────────────────

@app.route("/api/sessions/<session_id>/import-vi2", methods=["POST"])
def import_vi2(session_id):
    """上传一个 .vi2 文件，解析后导入为当前下一个未关联测点的数据"""
    if not VI2_AVAILABLE:
        return jsonify({"error": "缺少 vi2 解析库(olefile)，请在启动目录执行 pip install olefile"}), 500

    file = request.files.get("file")
    if not file:
        return jsonify({"error": "未收到文件"}), 400
    if not file.filename.lower().endswith(".vi2"):
        return jsonify({"error": "仅支持 .vi2 文件"}), 400

    # 保存临时文件
    tmp_dir = os.path.join(APP_DATA_DIR, "tmp")
    os.makedirs(tmp_dir, exist_ok=True)
    tmp_path = os.path.join(tmp_dir, f"{uuid.uuid4().hex}.vi2")
    file.save(tmp_path)

    # 需要文件首部是 OLE2 签名
    try:
        with open(tmp_path, "rb") as fh:
            sig = fh.read(8)
    except Exception as e:
        return jsonify({"error": f"读取文件失败: {e}"}), 500
    if sig != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return jsonify({"error": "不是有效的 .vi2 文件（缺少 OLE 签名）"}), 400

    # 解析
    sample_minutes = request.form.get("sample_minutes")
    try:
        sample_minutes = float(sample_minutes) if sample_minutes else None
    except ValueError:
        sample_minutes = None
    start_time = request.form.get("start_time") or None

    try:
        parsed = parse_vi2(tmp_path, sample_minutes=sample_minutes, start_time=start_time)
    except Exception as e:
        return jsonify({"error": f"解析 .vi2 失败: {e}"}), 500
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass

    if not parsed.get("records"):
        return jsonify({"error": "文件中未解析到温度数据"}), 400

    sn = parsed.get("serial_number", "未知")
    conn = get_db()
    next_point = conn.execute(
        "SELECT * FROM measurement_points WHERE session_id=? AND (serial_number='' OR serial_number IS NULL) "
        "ORDER BY sort_order LIMIT 1",
        (session_id,)
    ).fetchone()

    if not next_point:
        conn.close()
        return jsonify({"error": "所有测点都已关联设备，请新建会话或增加测点"}), 400

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for idx, rec in enumerate(parsed["records"], 1):
        conn.execute(
            "INSERT INTO device_records "
            "(session_id, point_number, serial_number, device_name, csv_file, "
            " record_index, date_val, time_val, temperature, humidity, alarm, raw_data, created_at) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (session_id, next_point["point_number"], sn, f"vi2-{sn}",
             "import.vi2", idx,
             rec["date"], rec["time"], rec["temperature"], None, "",
             json.dumps({"t_code": rec["t_code"], "unit": parsed.get("unit", "°C")}, ensure_ascii=False),
             now)
        )

    conn.execute(
        "UPDATE measurement_points SET serial_number=? WHERE id=?",
        (sn, next_point["id"])
    )
    completed = conn.execute(
        "SELECT COUNT(*) as cnt FROM measurement_points WHERE session_id=? AND serial_number!='' AND serial_number IS NOT NULL",
        (session_id,)
    ).fetchone()["cnt"]
    conn.execute("UPDATE sessions SET completed_points=? WHERE id=?", (completed, session_id))
    conn.commit()
    conn.close()

    return jsonify({
        "ok": True,
        "point_number": next_point["point_number"],
        "serial_number": sn,
        "record_count": len(parsed["records"]),
        "sample_minutes": parsed["sample_minutes"],
        "start_time": parsed["start_time"],
        "completed_points": completed,
        "message": f"测点 {next_point['point_number']} 导入 {len(parsed['records'])} 条数据 (SN: {sn})"
    })


# ─── API: 导出 Excel ─────────────────────────────────────────────────────────

@app.route("/api/sessions/<session_id>/export", methods=["POST"])
def export_data(session_id):
    data = request.get_json(silent=True) or {}
    selected = data.get("selected_indexes")
    filepath, filename = export_excel(session_id, selected)
    if filepath is None:
        return jsonify({"error": filename}), 400
    return send_file(filepath, as_attachment=True, download_name=filename)


# ─── API: 获取测点的数据预览 ──────────────────────────────────────────────────

@app.route("/api/sessions/<session_id>/points/<point_number>/data", methods=["GET"])
def get_point_data(session_id, point_number):
    conn = get_db()
    records = conn.execute(
        "SELECT * FROM device_records WHERE session_id=? AND point_number=? ORDER BY record_index LIMIT 100",
        (session_id, point_number)
    ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in records])


# ─── 启动 ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    init_db()
    print()
    print("╔══════════════════════════════════════════════════╗")
    print("║   Testo 184 温度计数据读取汇总工具              ║")
    print("╠══════════════════════════════════════════════════╣")
    print("║                                                  ║")
    print("║   请在浏览器中打开:                              ║")
    print("║   ➜  http://localhost:8000                       ║")
    print("║                                                  ║")
    print("║   按 Ctrl+C 停止服务                             ║")
    print("║                                                  ║")
    print("╚══════════════════════════════════════════════════╝")
    print()
    # 注意：不用 5000 端口，macOS 的 AirPlay 接收器默认占用 5000 端口

    # 自动在默认浏览器中打开（服务启动后延迟打开，避免端口未就绪）
    def _open_browser():
        import time
        time.sleep(1.0)
        import webbrowser
        try:
            webbrowser.open("http://localhost:8000")
        except Exception:
            pass

    threading.Thread(target=_open_browser, daemon=True).start()
    app.run(host="0.0.0.0", port=int(os.environ.get("TESTO_PORT", 8000)), debug=False)
