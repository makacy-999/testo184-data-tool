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

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("❌ 缺少 openpyxl，请运行: pip install openpyxl")
    sys.exit(1)


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
        name = os.path.basename(path).upper()
        if "TESTO" in name and "184" in name:
            return True
        try:
            files = [f.lower() for f in os.listdir(path)]
            testo_keywords = ["testo", "configuration", "measurement", "messdaten"]
            return sum(1 for kw in testo_keywords if any(kw in f for f in files)) >= 2
        except (OSError, PermissionError):
            return False

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
                try:
                    for root, dirs, files in os.walk(mp):
                        dirs[:] = [d for d in dirs if not d.startswith(".")]
                        for f in files:
                            if f.lower().endswith(".csv"):
                                csv_files.append(os.path.join(root, f))
                except (OSError, PermissionError) as e:
                    device["error"] = f"读取目录失败: {e}"

                device["csv_files"] = csv_files
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

                device["records"] = all_records
                # 尝试提取序列号
                for key in ["serial", "serial_number", "sn", "s/n", "序列号"]:
                    if key in device_info and device_info[key]:
                        device["serial_number"] = str(device_info[key])
                        break
                # 如果没找到，用设备名作为标识
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


# ─── Excel 导出 ─────────────────────────────────────────────────────────────

def export_excel(session_id):
    """将指定会话的数据导出为 Excel"""
    conn = get_db()
    session = conn.execute("SELECT * FROM sessions WHERE id=?", (session_id,)).fetchone()
    if not session:
        conn.close()
        return None, "会话不存在"

    points = conn.execute(
        "SELECT * FROM measurement_points WHERE session_id=? ORDER BY sort_order",
        (session_id,)
    ).fetchall()

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


# ─── API: 导出 Excel ─────────────────────────────────────────────────────────

@app.route("/api/sessions/<session_id>/export", methods=["POST"])
def export_data(session_id):
    filepath, filename = export_excel(session_id)
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
    app.run(host="0.0.0.0", port=8000, debug=False)
