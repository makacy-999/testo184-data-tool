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
import time
import base64
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
                xml_files = []
                try:
                    for root, dirs, files in os.walk(mp):
                        # 不排除隐藏目录：Testo 设备数据文件可能位于隐藏/系统目录（如 .SystemVolumeInformation）
                        dirs[:] = [d for d in dirs
                                   if d.lower() not in ("found.000", "system volume information", "已读")
                                   and d.lower() != "$recycle.bin"]
                        for f in files:
                            fl = f.lower()
                            if fl.endswith(".csv"):
                                csv_files.append(os.path.join(root, f))
                            elif fl.endswith((".xdp", ".xml")):
                                # Testo 184 设备盘的 "configuration_数据.xdp" 即 XML 数据包
                                xml_files.append(os.path.join(root, f))
                            elif fl.endswith(".vi2"):
                                vi2_files.append(os.path.join(root, f))
                            elif fl.endswith(".pdf"):
                                pdf_files.append(os.path.join(root, f))
                            else:
                                # 无扩展名/其他扩展名：按内容签名识别
                                fp = os.path.join(root, f)
                                try:
                                    with open(fp, "rb") as fh:
                                        sig = fh.read(2048)
                                    head = sig.lstrip(b"\xef\xbb\xbf \t\r\n")
                                    if head.startswith(b"<?xml") or head.startswith(b"<xdp"):
                                        xml_files.append(fp)
                                    elif head.startswith(b"%PDF"):
                                        pdf_files.append(fp)
                                    elif sig.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
                                        vi2_files.append(fp)
                                except Exception:
                                    pass
                except (OSError, PermissionError) as e:
                    device["error"] = f"读取目录失败: {e}"

                device["csv_files"] = csv_files
                device["pdf_files"] = [os.path.basename(f) for f in pdf_files]
                device["vi2_files"] = [os.path.basename(f) for f in vi2_files]
                device["xml_files"] = [os.path.basename(f) for f in xml_files]
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

                scan_log = []

                # XML / XDP 数据文件（设备盘上的 "configuration_数据.xdp" 等）
                if not all_records:
                    for xml_file in xml_files:
                        try:
                            headers, records, info = cls.parse_testo_xml(xml_file)
                            for k, v in info.items():
                                if v and not device_info.get(k):
                                    device_info[k] = v
                            scan_log.append({
                                "file": os.path.basename(xml_file), "kind": "XML/XDP",
                                "status": "ok" if records else "empty",
                                "records": len(records),
                                "detail": ("提取 %d 条 measurement 记录" % len(records) if records
                                           else "文件中没有 measurement/record 数据记录"),
                            })
                            if records:
                                for r in records:
                                    r["source_file"] = os.path.basename(xml_file)
                                all_records.extend(records)
                        except Exception as e:
                            scan_log.append({"file": os.path.basename(xml_file), "kind": "XML/XDP",
                                             "status": "error", "records": 0, "detail": str(e)})
                            device["error"] = "解析 %s 失败: %s" % (os.path.basename(xml_file), e)

                # 如果没有 CSV/XML，尝试从 Testo 报告 PDF 中提取数据（内置报告）
                # 注：图形化报告即使提不出曲线数据，其文本里的 SN 也要保留
                if not all_records:
                    for pdf_file in pdf_files:
                        try:
                            headers, records, info = cls.parse_testo_pdf(pdf_file)
                            for k, v in info.items():
                                if v and not device_info.get(k):
                                    device_info[k] = v
                            scan_log.append({
                                "file": os.path.basename(pdf_file), "kind": "PDF",
                                "status": "ok" if records else "empty",
                                "records": len(records),
                                "detail": (info.get("chart_error") or
                                           ("曲线提取 %d 点" % len(records) if info.get("source") == "pdf_chart"
                                            else ("表格提取 %d 行" % len(records) if records else "无数据表格/曲线"))),
                            })
                            if records:
                                for r in records:
                                    r["source_file"] = os.path.basename(pdf_file)
                                all_records.extend(records)
                                break
                        except Exception as e:
                            scan_log.append({"file": os.path.basename(pdf_file), "kind": "PDF",
                                             "status": "error", "records": 0, "detail": str(e)})
                            device["error"] = f"解析 {os.path.basename(pdf_file)} 失败: {e}"

                # 若仍无数据，尝试解析 .vi2 专有存档（Testo ComSoft 导出格式）
                if not all_records and VI2_AVAILABLE:
                    for vi2_file in vi2_files:
                        try:
                            parsed = parse_vi2(vi2_file)
                            scan_log.append({
                                "file": os.path.basename(vi2_file), "kind": "vi2",
                                "status": "ok" if parsed.get("records") else "empty",
                                "records": len(parsed.get("records") or []),
                                "detail": ("OLE 存档提取 %d 条" % len(parsed.get("records") or [])
                                           if parsed.get("records") else "存档中没有数据记录"),
                            })
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
                device["scan_log"] = scan_log
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
                # 从报告/校准证书/xdp 的文本中挖掘 SN（PDF 页眉页脚、证书正文都印有 SN）
                if not device["serial_number"]:
                    for f in pdf_files + xml_files:
                        try:
                            if f.lower().endswith((".xml", ".xdp")):
                                txt = open(f, "rb").read(65536).decode("utf-8", "ignore")
                            else:
                                with pdfplumber.open(f) as pdf:
                                    txt = "\n".join((pg.extract_text() or "") for pg in pdf.pages[:3])
                            sn = DeviceDetector._extract_sn_from_text(txt)
                            if sn:
                                device["serial_number"] = sn
                                break
                        except Exception:
                            continue
                # 如果实在提取不到，拼接设备名作为标识（用户可在界面手动修正）
                if not device["serial_number"]:
                    device["serial_number"] = device["name"]

                devices.append(device)

        # ── v3.0.1: 追加扫描 ComSoft 接力文件夹中的全量 vi2/CSV 存档 ──
        # testo 184 设备 U 盘上没有原始数据文件，只有自带 PDF 报告（官方曲线最多 324 点）。
        # 全量数据须由 ComSoft(cc4.exe) 读取设备后导出为 .vi2/.csv 到接力文件夹(testo_export)。
        # 这里把接力文件夹中的存档也识别为候选设备源，供批量流程读取全量数据。
        try:
            relay_dirs = cls._relay_export_dirs()
            seen_paths = {d.get("path") for d in devices}
            for rdir in relay_dirs:
                for root, dirs, files in os.walk(rdir):
                    dirs[:] = [d for d in dirs
                               if d.lower() not in ("found.000", "system volume information", "已读")
                               and d.lower() != "$recycle.bin"]
                    for f in files:
                        fl = f.lower()
                        if not (fl.endswith(".vi2") or fl.endswith(".csv")):
                            continue
                        full = os.path.join(root, f)
                        try:
                            if fl.endswith(".csv"):
                                headers, records, rinfo = cls.parse_csv(full)
                                for r in records:
                                    r["source_file"] = os.path.basename(f)
                            else:
                                parsed = parse_vi2(full)
                                records = []
                                for idx, r in enumerate(parsed.get("records") or [], 1):
                                    records.append({"date": r["date"], "time": r["time"],
                                                    "temperature": r["temperature"],
                                                    "source_file": os.path.basename(f),
                                                    "record_index": idx})
                                rinfo = {"serial": parsed.get("serial_number", ""),
                                     "unit": parsed.get("unit", "°C"),
                                     "limit_min": parsed.get("limit_min"),
                                     "limit_max": parsed.get("limit_max"),
                                     "start_time": parsed.get("start_time", "")}
                            if not records:
                                continue
                            sn = str((rinfo or {}).get("serial") or "")
                            if not sn:
                                import re as _re
                                m = _re.search(r"(\d{6,})", os.path.basename(f))
                                if m:
                                    sn = m.group(1)
                            if not sn:
                                sn = "RELAY_" + f
                            dev = {
                                "path": full,
                                "name": os.path.basename(f),
                                "csv_files": [full] if fl.endswith(".csv") else [],
                                "vi2_files": [os.path.basename(f)] if fl.endswith(".vi2") else [],
                                "pdf_files": [],
                                "xml_files": [],
                                "records": records,
                                "serial_number": sn,
                                "unit": (rinfo or {}).get("unit", "°C"),
                                "limit_min": (rinfo or {}).get("limit_min"),
                                "limit_max": (rinfo or {}).get("limit_max"),
                                "start_time": (rinfo or {}).get("start_time", ""),
                                "scan_log": [{"file": os.path.basename(f), "kind": "ComSoft存档",
                                              "status": "ok", "records": len(records),
                                              "detail": "接力文件夹全量存档 %d 条" % len(records)}],
                                "error": None,
                                "_relay": True,
                            }
                            from itertools import count
                            _c = count(1)
                            key = full
                            probes = [key] + [key + str(next(_c)) for _ in range(5)]
                            if key not in seen_paths:
                                devices.append(dev)
                                seen_paths.add(key)
                        except Exception:
                            continue
        except Exception:
            pass
        return devices

    @classmethod
    def _relay_export_dirs(cls):
        """返回需要扫描的 ComSoft 接力导出目录（去重后）"""
        dirs = []
        try:
            dirs.append(_recommended_export_dir())
        except Exception:
            pass
        home = os.path.expanduser("~")
        try:
            raw = _ensure_desktop_raw_folder()
            if raw:
                dirs.append(raw)
        except Exception:
            pass
        for base in ("Desktop", "桌面", "Documents", "文档",
                     "OneDrive/Desktop", "OneDrive/桌面", "OneDrive/文档"):
            d = os.path.join(home, base)
            if os.path.isdir(d):
                dirs.append(d)
        seen = set()
        out = []
        for d in dirs:
            try:
                real = os.path.realpath(d)
            except OSError:
                real = d
            if real not in seen:
                seen.add(real)
                out.append(d)
        return out

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

        # 自动检测分隔符：Sniffer 对单列/纯数值行常失败，退回统计法
        sample = "\n".join(data_lines[:10])
        sep = ","
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            if dialect.delimiter:
                sep = dialect.delimiter
        except csv.Error:
            pass
        if sep == ",":
            best = None
            best_score = -1
            for cand in [",", ";", "\t", "|"]:
                try:
                    cnt = [len(r) for r in csv.reader(data_lines[:20], delimiter=cand) if r]
                except Exception:
                    continue
                if not cnt:
                    continue
                # 一致性：>60% 行分割后列数相等；且列数 >=2
                from collections import Counter
                top, freq = Counter(cnt).most_common(1)[0]
                score = freq / len(cnt)
                if score > best_score and top >= 2:
                    best_score = score
                    best = (cand, top, score)
            if best and best[2] >= 0.6:
                sep = best[0]

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
    def _parse_temperature_text(txt):
        """'4,5 °C' -> 4.5 ；无效返回 None"""
        if txt is None:
            return None
        s = str(txt).strip()
        if not s:
            return None
        import re as _re
        m = _re.search(r"(-?\d+(?:[.,]\d+)?)", s)
        if not m:
            return None
        try:
            return float(m.group(1).replace(",", "."))
        except ValueError:
            return None

    @staticmethod
    def parse_testo_xml(xml_path):
        """解析 Testo 导出的 XML / XDP 数据文件。

        典型来源：设备 U 盘上的 "testo 184 configuration_数据.xdp"
        （XDP = XML Data Package，本质是 XML）。

        参考结构（用户提供）：
            <root>
                <measurement>
                    <time>2026-09-20 15:00:00</time>
                    <temperature>4.2</temperature>
                </measurement>
                ...
            </root>

        兼容：XML 命名空间、标签/属性变体、多种时间格式
        （ISO / 中式 / 美式 / 德式 / epoch 秒毫秒）、逗号小数、
        °C 单位残留，humidity 一并提取。
        """
        import xml.etree.ElementTree as ET

        with open(xml_path, "rb") as fh:
            raw = fh.read()
        try:
            root = ET.fromstring(raw)
        except ET.ParseError:
            head = raw.lstrip(b"\xef\xbb\xbf \t\r\n")
            try:
                root = ET.fromstring(head)
            except ET.ParseError as e2:
                raise ValueError("XML 解析失败: %s" % e2)

        def local(tag):
            return tag.rsplit("}", 1)[-1].strip().lower()

        TIME_TAGS = {"time", "datetime", "timestamp", "date", "datum", "zeit",
                     "meastime", "measurementtime", "recordtime"}
        TEMP_TAGS = {"temperature", "temperatur", "temp", "t", "value",
                     "measvalue", "measuredvalue"}
        HUM_TAGS = {"humidity", "hum", "rhumidity", "relativehumidity", "feuchte"}
        CONTAINERS = {"measurement", "record", "reading", "row", "datapoint",
                      "sample", "logvalue", "entry"}

        def field_text(container, tags, depth_limit=3):
            """在容器自身属性及下 depth_limit 层内（BFS）找标签/属性匹配的值"""
            # 容器自身的属性：<measurement time="..." temperature="..."/>
            for k, v in container.attrib.items():
                if local(k) in tags and v.strip():
                    return v.strip()
            queue = [(child, 1) for child in container]
            while queue:
                cur, d = queue.pop(0)
                ln = local(cur.tag)
                if ln in tags:
                    txt = (cur.text or "").strip() if cur.text else ""
                    if txt:
                        return txt
                for k, v in cur.attrib.items():
                    if local(k) in tags and v.strip():
                        return v.strip()
                if d < depth_limit:
                    queue.extend((c, d + 1) for c in cur)

        def has_nested_container(el, depth_limit=6):
            stack = [(child, 1) for child in el]
            while stack:
                cur, d = stack.pop()
                if local(cur.tag) in CONTAINERS:
                    return True
                if d < depth_limit:
                    stack.extend((c, d + 1) for c in cur)
            return False

        records = []
        for c in root.iter():
            if local(c.tag) not in CONTAINERS:
                continue
            if has_nested_container(c):
                continue  # 外层容器跳过，只解析最内层，避免重复计数
            time_txt = field_text(c, TIME_TAGS)
            temp_txt = field_text(c, TEMP_TAGS)
            if time_txt is None and temp_txt is None:
                continue
            hum_txt = field_text(c, HUM_TAGS)
            records.append({
                "date": "", "time": "",
                "temperature": DeviceDetector._parse_temperature_text(temp_txt),
                "humidity": DeviceDetector._parse_temperature_text(hum_txt),
                "alarm": "",
                "_raw": {"time_text": time_txt or "", "temp_text": temp_txt or ""},
            })

        # 时间文本 -> date / time 字段
        import re as _re
        from datetime import datetime as _dtm
        TIME_FORMATS = [
            "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
            "%Y-%m-%dT%H:%M", "%Y.%m.%d %H:%M:%S", "%Y.%m.%d %H:%M",
            "%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M",
            "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M",
            "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M",
        ]
        for r in records:
            t = (r["_raw"].get("time_text") or "").strip()
            d, tm = "", ""
            if t:
                parsed = None
                if _re.fullmatch(r"\d{9,13}", t):  # epoch 秒 / 毫秒
                    try:
                        ts = int(t)
                        if ts > 10 ** 12:
                            ts /= 1000.0
                        parsed = _dtm.fromtimestamp(ts)
                    except (ValueError, OSError, OverflowError):
                        parsed = None
                else:
                    for fmt in TIME_FORMATS:
                        try:
                            parsed = _dtm.strptime(t, fmt)
                            break
                        except ValueError:
                            continue
                if parsed is None:
                    if _re.fullmatch(r"\d{1,2}:\d{2}(:\d{2})?", t):
                        tm = t if t.count(":") == 2 else t + ":00"
                    else:
                        d, tm = t, ""  # 保留原文，避免丢数据
                else:
                    d = parsed.strftime("%Y-%m-%d")
                    tm = parsed.strftime("%H:%M:%S")
            r["date"], r["time"] = d, tm

        # 温度与时间都拿不到的记录丢弃
        records = [r for r in records
                   if r["temperature"] is not None or (r["date"] or r["time"])]

        # 提取序列号：任意元素/属性名含 serial
        info = {"serial_number": "", "xml": True, "source": "xml"}
        sn_found = ""
        for el in root.iter():
            if "serial" in local(el.tag) and (el.text or "").strip():
                sn_found = el.text.strip()
                break
            for k, v in el.attrib.items():
                if "serial" in local(k) and v.strip():
                    sn_found = v.strip()
                    break
            if sn_found:
                break
        info["serial_number"] = sn_found

        headers = ["date", "time", "temperature", "humidity", "alarm"]
        return headers, records, info

    @staticmethod
    def _extract_sn_from_text(text):
        """从任意文本中提取 Testo 设备序列号（Serial Number / S/N / Seriennummer）"""
        if not text:
            return ""
        import re as _re
        pats = [
            r"(?:serial(?:\s*number)?|s\s*/\s*n|seriennummer|serial\s*no\.?|序列号)"
            r"[^0-9A-Za-z]{0,4}[:：]?\s*([0-9]{6,10})",
            r"\bSN\.?\s*[:：.]?\s*([0-9]{6,10})",   # 如 "SN.:44023840"（testo 184 报告页脚）
            r"\b(440[0-9]{5})\b",   # testo 184 SN 特征段
        ]
        for pat in pats:
            m = _re.search(pat, text, _re.IGNORECASE)
            if m:
                return m.group(1)
        return ""

    @staticmethod
    def parse_pdf_chart(pdf_path):
        """从 Testo 图形化 PDF 报告的矢量曲线中提取温度数据。

        设备 U 盘自动生成的 "measurement report" 是曲线图 PDF：没有数据表格，
        数据藏在矢量路径坐标里。原理：
        - 曲线点的 y 坐标经左缘温度刻度（等差数值文本）线性映射为温度
        - x 坐标经底部时间刻度线性映射为时间
        - 水平/垂直长线段（网格、边框、限值线）被过滤
        返回 (records, info)。
        """
        import re as _re
        from datetime import datetime as _dtm, timedelta as _td

        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            words = page.extract_words() or []

            # ── 温度刻度：左缘数值词按 x 聚类，找单调等差的一列 ──
            nums = []
            for w in words:
                t = w["text"].strip().replace(",", ".")
                if _re.fullmatch(r"-?\d+(?:\.\d+)?", t):
                    try:
                        v = float(t)
                    except ValueError:
                        continue
                    nums.append({"v": v, "x": (w["x0"] + w["x1"]) / 2,
                                 "y": (w["top"] + w["bottom"]) / 2})
            temp_axis = None
            buckets = {}
            for n in nums:
                buckets.setdefault(round(n["x"] / 6), []).append(n)
            for _, grp in sorted(buckets.items(), key=lambda kv: kv[1][0]["x"]):
                if len(grp) < 3:
                    continue
                grp.sort(key=lambda n: n["y"])
                vals = [n["v"] for n in grp]
                mono = (all(vals[i] < vals[i + 1] for i in range(len(vals) - 1))
                        or all(vals[i] > vals[i + 1] for i in range(len(vals) - 1)))
                if not mono:
                    continue
                diffs = [abs(vals[i + 1] - vals[i]) for i in range(len(vals) - 1)]
                base = max(diffs)
                if base <= 0 or min(diffs) < base * 0.7:
                    continue
                temp_axis = grp  # 最左的合格刻度列
                break
            if temp_axis is None or len(temp_axis) < 2:
                raise ValueError("未找到温度坐标轴刻度")
            y1, v1 = temp_axis[0]["y"], temp_axis[0]["v"]
            y2, v2 = temp_axis[-1]["y"], temp_axis[-1]["v"]
            if abs(y2 - y1) < 1:
                raise ValueError("温度刻度异常")
            t_slope = (v2 - v1) / (y2 - y1)

            # ── 时间刻度：最底部一行的 HH:MM 文本 ──
            time_words = []
            for w in words:
                t = w["text"].strip()
                m = _re.fullmatch(r"(\d{1,2}):(\d{2})(?::(\d{2}))?", t)
                if m:
                    sec = int(m.group(1)) * 3600 + int(m.group(2)) * 60 + int(m.group(3) or 0)
                    time_words.append({"sec": sec, "x": w["x0"],
                                       "y": (w["top"] + w["bottom"]) / 2})
            x_slope = None
            if len(time_words) >= 2:
                time_words.sort(key=lambda w: -w["y"])
                cand = [w for w in time_words if abs(w["y"] - time_words[0]["y"]) < 12]
                cand.sort(key=lambda w: w["x"])
                # 等差链清洗：统计区的起止时间等游离词可能与刻度行同一水平，
                # 且时间值恰好衔接；要求相邻刻度的时间差与 x 间距同时均匀，
                # 只保留最长等差链
                diffs = [cand[i + 1]["sec"] - cand[i]["sec"] for i in range(len(cand) - 1)]
                dxs = [cand[i + 1]["x"] - cand[i]["x"] for i in range(len(cand) - 1)]
                pos = sorted(d for d in diffs if d > 0)
                posx = sorted(d for d in dxs if d > 0)
                med = pos[len(pos) // 2] if pos else 0
                medx = posx[len(posx) // 2] if posx else 0
                if med > 0 and medx > 0:
                    best = []
                    for s in range(len(cand)):
                        chain = [cand[s]]
                        for w in cand[s + 1:]:
                            d = w["sec"] - chain[-1]["sec"]
                            dx = w["x"] - chain[-1]["x"]
                            if (abs(d - med) <= med * 0.3
                                    and abs(dx - medx) <= medx * 0.45):
                                chain.append(w)
                        if len(chain) > len(best):
                            best = chain
                    if len(best) >= 2:
                        cand = best
                if len(cand) >= 2 and cand[-1]["x"] - cand[0]["x"] > 1:
                    x_slope = (cand[-1]["sec"] - cand[0]["sec"]) / (cand[-1]["x"] - cand[0]["x"])
                    x_base_x, x_base_sec = cand[0]["x"], cand[0]["sec"]

            # 报告日期（时间轴只有时分时作基准日）
            page_text = page.extract_text() or ""
            base_date = ""
            m = _re.search(r"\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}\.\d{1,2}\.\d{4}", page_text)
            if m:
                s = m.group(0)
                for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%d.%m.%Y"):
                    try:
                        base_date = _dtm.strptime(s, fmt).strftime("%Y-%m-%d")
                        break
                    except ValueError:
                        continue

            # ── 曲线点：只取 curves 连续折线（testo 报告的数据曲线是单条 polyline），
            # lines 是网格/边框/限值线/刻度短线，整体排除；并用图表区域双重过滤 ──
            # （单段可能近似水平，不能按段过滤；区域外可能有 logo 等装饰曲线）
            ax_ys = [n["y"] for n in temp_axis]
            y_lo, y_hi = min(ax_ys) - 15, max(ax_ys) + 15
            if x_slope is not None:
                x_lo, x_hi = x_base_x - 20, cand[-1]["x"] + 20
            else:
                x_lo, x_hi = temp_axis[0]["x"] + 5, page.width
            pts = []
            for obj in list(page.curves):
                seg = obj.get("pts") or []
                for q in seg:
                    qx, qy = q[0], q[1]
                    if x_lo <= qx <= x_hi and y_lo <= qy <= y_hi:
                        pts.append((qx, qy))
            if len(pts) < 10:
                raise ValueError("未找到数据曲线")
            pts.sort(key=lambda q: q[0])
            merged = []
            for x, y in pts:
                if merged and x - merged[-1][0] < 0.7:
                    px, py, k = merged[-1]
                    merged[-1] = [px, (py * k + y) / (k + 1), k + 1]
                else:
                    merged.append([x, y, 1])

            records = []
            for x, y, _k in merged:
                temp = v1 + t_slope * (y - y1)
                rec = {"date": "", "time": "", "temperature": round(temp, 2),
                       "humidity": None, "alarm": "",
                       "_raw": {"source": "chart", "x": round(x, 1), "y": round(y, 1)}}
                if x_slope is not None:
                    sec = x_base_sec + x_slope * (x - x_base_x)
                    sec = int(round(sec))
                    days, rem = divmod(sec, 86400)
                    hh, rem2 = divmod(rem, 3600)
                    mm, ss = divmod(rem2, 60)
                    if base_date:
                        try:
                            d0 = _dtm.strptime(base_date, "%Y-%m-%d") + _td(days=days)
                            rec["date"] = d0.strftime("%Y-%m-%d")
                        except ValueError:
                            pass
                    rec["time"] = "%02d:%02d:%02d" % (hh, mm, ss)
                records.append(rec)
            info = {"source": "pdf_chart", "chart_points": len(records)}
            return records, info

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

        # 回退：矢量曲线图报告（设备 U 盘自动生成的报告没有数据表格）
        if not records:
            try:
                chart_records, chart_info = DeviceDetector.parse_pdf_chart(pdf_path)
                records = chart_records
                device_info.update({k: v for k, v in chart_info.items() if v})
            except Exception as e:
                device_info["chart_error"] = str(e)

        device_info["pdf_file"] = os.path.basename(pdf_path)
        if not device_info.get("source"):
            device_info["source"] = "pdf_report"
        # 从全文提取设备 SN（报告页眉/页脚通常印有 Serial Number）
        try:
            with pdfplumber.open(pdf_path) as pdf:
                full_text = "\n".join((pg.extract_text() or "") for pg in pdf.pages[:5])
        except Exception:
            full_text = ""
        sn = DeviceDetector._extract_sn_from_text(full_text)
        if sn:
            device_info["serial_number"] = sn
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
                    "scan_log": dev.get("scan_log", []),
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


# ─── vi2 导入核心（上传路由与目录监控共用） ──────────────────────────────────

def _looks_like_csv(path, head):
    """按扩展名或文本特征判断是否按 CSV 解析（ComSoft 导出常用逗号/分号分隔）"""
    if os.path.splitext(path)[1].lower() == ".csv":
        return True
    if not head or head.startswith(b"\xd0\xcf") or head.startswith(b"%PDF"):
        return False
    if b"\n" not in head:
        return False
    if b"," not in head and b";" not in head and b"\t" not in head:
        return False
    return all(32 <= b < 127 or b in (9, 10, 13) or b >= 160 for b in head[:256])


def _import_data_to_session(session_id, path, sample_minutes=None, start_time=None):
    """把数据文件（.vi2 存档 / .xml·.xdp 数据包）解析并导入到会话的下一个未关联测点。
    返回 (ok: bool, payload: dict)"""
    try:
        with open(path, "rb") as fh:
            sig = fh.read(2048)
    except Exception as e:
        return False, {"error": f"读取文件失败: {e}"}
    head = sig.lstrip(b"\xef\xbb\xbf \t\r\n")
    is_ole = sig.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")
    is_xml = (os.path.splitext(path)[1].lower() in (".xml", ".xdp")
              or head.startswith(b"<?xml") or head.startswith(b"<xdp"))
    if is_ole:
        if not VI2_AVAILABLE:
            return False, {"error": "缺少 vi2 解析库(olefile)，请执行 pip install olefile"}
        try:
            parsed = parse_vi2(path, sample_minutes=sample_minutes, start_time=start_time)
        except Exception as e:
            return False, {"error": f"解析 .vi2 失败: {e}"}
        if not parsed.get("records"):
            return False, {"error": "文件中未解析到温度数据"}
        sn = parsed.get("serial_number", "未知")
        records = [{"date": r["date"], "time": r["time"],
                    "temperature": r["temperature"], "humidity": None,
                    "t_code": r["t_code"]} for r in parsed["records"]]
        device_label = f"vi2-{sn}"
        raw_unit = parsed.get("unit", "°C")
    elif is_xml:
        try:
            _headers, records, info = DeviceDetector.parse_testo_xml(path)
        except Exception as e:
            return False, {"error": f"解析 XML/XDP 失败: {e}"}
        if not records:
            return False, {"error": "XML 文件中未解析到温度数据（没有 measurement/record 记录）"}
        sn = info.get("serial_number") or "未知"
        device_label = f"xml-{sn}"
        raw_unit = "°C"
    elif _looks_like_csv(path, head):
        # ComSoft 专业版导出的 CSV（分隔符逗号/分号，德文/中文/英文表头均可）
        try:
            _headers, records, info = DeviceDetector.parse_csv(path)
        except Exception as e:
            return False, {"error": f"解析 CSV 失败: {e}"}
        if not records:
            return False, {"error": "CSV 文件中未解析到温度数据"}
        sn = info.get("serial_number") or "未知"
        if sn == "未知":
            try:
                with open(path, "r", encoding="utf-8", errors="ignore") as fh:
                    sn = DeviceDetector._extract_sn_from_text(fh.read(4096)) or "未知"
            except OSError:
                pass
        device_label = f"csv-{sn}"
        raw_unit = "°C"
        records = [{"date": r.get("date", ""), "time": r.get("time", ""),
                    "temperature": DeviceDetector._parse_temperature_text(r.get("temperature")),
                    "humidity": DeviceDetector._parse_temperature_text(r.get("humidity")),
                    "_raw": r.get("_raw")} for r in records]
    elif head.startswith(b"%PDF") or os.path.splitext(path)[1].lower() == ".pdf":
        # Testo 报告 PDF：数据表格或图形曲线（设备盘 measurement report）
        try:
            _headers, pdf_records, info = DeviceDetector.parse_testo_pdf(path)
        except Exception as e:
            return False, {"error": f"解析 PDF 失败: {e}"}
        if not pdf_records:
            return False, {"error": "PDF 中未解析到温度数据（表格与曲线提取均无结果）"}
        sn = info.get("serial_number") or "未知"
        device_label = f"pdf-{sn}"
        raw_unit = "°C"
        records = [{"date": r.get("date", ""), "time": r.get("time", ""),
                    "temperature": DeviceDetector._parse_temperature_text(r.get("temperature")),
                    "humidity": DeviceDetector._parse_temperature_text(r.get("humidity")),
                    "_raw": r.get("_raw")} for r in pdf_records]
    else:
        return False, {"error": "不是支持的数据文件（支持 .vi2 存档、CSV 导出、.xml / .xdp 数据包、Testo 报告 PDF）"}
    conn = get_db()
    next_point = conn.execute(
        "SELECT * FROM measurement_points WHERE session_id=? AND (serial_number='' OR serial_number IS NULL) "
        "ORDER BY sort_order LIMIT 1",
        (session_id,)
    ).fetchone()
    if not next_point:
        conn.close()
        return False, {"error": "所有测点都已关联设备，请增加测点或新建会话"}
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for idx, rec in enumerate(records, 1):
        raw = {"unit": raw_unit}
        if rec.get("t_code") is not None:
            raw["t_code"] = rec["t_code"]
        if rec.get("_raw"):
            raw.update(rec["_raw"])
        conn.execute(
            "INSERT INTO device_records "
            "(session_id, point_number, serial_number, device_name, csv_file, "
            " record_index, date_val, time_val, temperature, humidity, alarm, raw_data, created_at) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (session_id, next_point["point_number"], sn, device_label,
             os.path.basename(path), idx,
             rec["date"], rec["time"], rec["temperature"], rec.get("humidity"), "",
             json.dumps(raw, ensure_ascii=False),
             now)
        )
    conn.execute("UPDATE measurement_points SET serial_number=? WHERE id=?", (sn, next_point["id"]))
    completed = conn.execute(
        "SELECT COUNT(*) as cnt FROM measurement_points WHERE session_id=? AND serial_number!='' AND serial_number IS NOT NULL",
        (session_id,)
    ).fetchone()["cnt"]
    conn.execute("UPDATE sessions SET completed_points=? WHERE id=?", (completed, session_id))
    conn.commit()
    conn.close()
    vi2_meta = parsed if is_ole else {}
    return True, {
        "point_number": next_point["point_number"],
        "serial_number": sn,
        "record_count": len(records),
        "sample_minutes": vi2_meta.get("sample_minutes"),
        "start_time": vi2_meta.get("start_time"),
        "completed_points": completed,
        "message": f"测点 {next_point['point_number']} 导入 {len(records)} 条数据 (SN: {sn})"
    }


# ─── API: 导入 .vi2 文件（支持一次选择多个） ─────────────────────────────────

@app.route("/api/sessions/<session_id>/import-vi2", methods=["POST"])
def import_vi2(session_id):
    """上传一个或多个数据文件（.vi2 / .xml / .xdp），依次导入到未关联测点"""
    files = request.files.getlist("file")
    if not files:
        return jsonify({"error": "未收到文件"}), 400
    tmp_dir = os.path.join(APP_DATA_DIR, "tmp")
    os.makedirs(tmp_dir, exist_ok=True)
    sample_minutes = request.form.get("sample_minutes")
    try:
        sample_minutes = float(sample_minutes) if sample_minutes else None
    except (TypeError, ValueError):
        sample_minutes = None
    start_time = request.form.get("start_time") or None
    results = []
    for file in files:
        ext = os.path.splitext(file.filename or "")[1].lower()
        if not ext:
            ext = ".bin"
        tmp_path = os.path.join(tmp_dir, f"{uuid.uuid4().hex}{ext}")
        try:
            file.save(tmp_path)
            ok, payload = _import_data_to_session(session_id, tmp_path, sample_minutes, start_time)
            entry = dict(payload)
            entry["ok"] = ok
            entry["file"] = file.filename
            results.append(entry)
        finally:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
    ok_any = any(r.get("ok") for r in results)
    if ok_any:
        msg = "；".join(r.get("message") or r.get("error", "") for r in results)
    else:
        msg = results[0].get("error", "导入失败") if results else "导入失败"
    return jsonify({"ok": ok_any, "results": results, "message": msg})


# ─── Testo / ComSoft 软件联动：监控导出目录，新 .vi2 自动导入转 Excel ─────────

WATCHERS = {}  # session_id -> {"path", "stop", "results", "files", "pending", "ready"}


def _list_data_files(path):
    """列出目录下可作为数据导入的文件（.vi2 / .xml / .xdp 及签名匹配的无扩展名文件）"""
    out = []
    try:
        for f in os.listdir(path):
            fp = os.path.join(path, f)
            if not os.path.isfile(fp):
                continue
            if f.lower().endswith((".vi2", ".xdp", ".xml", ".csv", ".pdf")):
                out.append(fp)
                continue
            # Testo 软件导出的文件可能无扩展名：按内容签名识别
            try:
                with open(fp, "rb") as fh:
                    sig = fh.read(2048)
                head = sig.lstrip(b"\xef\xbb\xbf \t\r\n")
                if (sig.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")
                        or head.startswith(b"<?xml") or head.startswith(b"<xdp")
                        or head.startswith(b"%PDF")):
                    out.append(fp)
            except OSError:
                pass
    except OSError:
        pass
    return out


def _watch_loop(session_id, path, stop_event):
    w = WATCHERS.get(session_id)
    if w is None:
        return
    # 初始快照：已存在的文件不重复导入
    snap = {}
    for f in _list_data_files(path):
        try:
            snap[f] = os.path.getsize(f)
        except OSError:
            pass
    w["files"] = snap
    w["pending"] = {}
    # 全量文件快照（诊断用：能看到 Testo 软件保存的任何新文件）
    try:
        w["all_files"] = {f: os.path.getsize(os.path.join(path, f))
                          for f in os.listdir(path)
                          if os.path.isfile(os.path.join(path, f))}
    except OSError:
        w["all_files"] = {}
    w["new_files_seen"] = []
    w["ready"] = True
    while not stop_event.wait(2.5):
        w = WATCHERS.get(session_id)
        if w is None:
            return
        try:
            # 全量快照：任何新出现的文件都记录（诊断）
            try:
                current_all = {f: os.path.getsize(os.path.join(path, f))
                               for f in os.listdir(path)
                               if os.path.isfile(os.path.join(path, f))}
            except OSError:
                current_all = {}
            for f, size in current_all.items():
                if f in w["all_files"]:
                    continue
                w["all_files"][f] = size
                if not f.lower().endswith((".vi2", ".xdp", ".xml", ".csv", ".pdf")):
                    # 不是数据格式的新文件 → 记入诊断（可能是 Testo 保存的其他格式）
                    seen = w["new_files_seen"]
                    if len(seen) < 30 and f not in [x.get("file") for x in seen]:
                        seen.append({"file": f, "size": size,
                                     "note": "新文件但不是数据格式（.vi2/.csv/.xml/.xdp/PDF/OLE2），未导入"})
            # 数据文件检测
            current = {}
            for f in _list_data_files(path):
                try:
                    current[f] = os.path.getsize(f)
                except OSError:
                    continue
            for f, size in current.items():
                if f in w["files"]:
                    continue
                if w["pending"].get(f) == size:
                    # 大小两轮一致，文件已写完 → 自动导入
                    ok, payload = _import_data_to_session(session_id, f)
                    entry = dict(payload)
                    entry["ok"] = ok
                    entry["file"] = os.path.basename(f)
                    w["results"].append(entry)
                    w["files"][f] = size
                    w["pending"].pop(f, None)
                else:
                    w["pending"][f] = size
            for f in list(w["pending"].keys()):
                if f not in current:
                    w["pending"].pop(f, None)
        except Exception:
            pass


def _stop_watcher(session_id):
    w = WATCHERS.pop(session_id, None)
    if w:
        w["stop"].set()


@app.route("/api/sessions/<session_id>/watch-folder", methods=["POST"])
def watch_folder(session_id):
    """启动目录监控：Testo/ComSoft 软件保存 .vi2 到该目录时自动导入"""
    data = request.get_json(silent=True) or {}
    path = (data.get("path") or "").strip().strip('"')
    if not path or not os.path.isdir(path):
        return jsonify({"error": f"目录不存在: {path or '(空)'}"}), 400
    _stop_watcher(session_id)
    stop = threading.Event()
    WATCHERS[session_id] = {
        "path": path, "stop": stop, "results": [],
        "files": {}, "pending": {}, "ready": False,
    }
    t = threading.Thread(target=_watch_loop, args=(session_id, path, stop), daemon=True)
    WATCHERS[session_id]["thread"] = t
    t.start()
    return jsonify({"ok": True, "watching": path})


@app.route("/api/sessions/<session_id>/watch-status", methods=["GET"])
def watch_status(session_id):
    w = WATCHERS.get(session_id)
    if not w:
        return jsonify({"watching": False, "new_imports": [], "new_files_seen": []})
    results = w["results"]
    w["results"] = []
    seen = w.get("new_files_seen", [])
    w["new_files_seen"] = []
    return jsonify({"watching": True, "path": w["path"],
                    "ready": w.get("ready", False), "new_imports": results,
                    "new_files_seen": seen})


@app.route("/api/sessions/<session_id>/watch-stop", methods=["POST"])
def watch_stop(session_id):
    _stop_watcher(session_id)
    return jsonify({"ok": True})


def _find_testo_software():
    """Windows 上在注册表查找已安装的 Testo / ComSoft 软件"""
    results = []
    if sys.platform != "win32":
        return results
    try:
        import winreg
        seen = set()
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            for wow in (winreg.KEY_WOW64_32KEY, winreg.KEY_WOW64_64KEY, 0):
                try:
                    key = winreg.OpenKey(root, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                                         0, winreg.KEY_READ | wow)
                except OSError:
                    continue
                i = 0
                while True:
                    try:
                        sub = winreg.EnumKey(key, i)
                    except OSError:
                        break
                    i += 1
                    try:
                        sk = winreg.OpenKey(key, sub)
                        try:
                            name = str(winreg.QueryValueEx(sk, "DisplayName")[0])
                        except OSError:
                            continue
                        low = name.lower()
                        if ("comsoft" in low or "testo" in low) and name not in seen:
                            try:
                                loc = str(winreg.QueryValueEx(sk, "InstallLocation")[0] or "")
                            except OSError:
                                loc = ""
                            if not loc:
                                try:
                                    loc = str(winreg.QueryValueEx(sk, "InstallSource")[0] or "")
                                except OSError:
                                    loc = ""
                            seen.add(name)
                            results.append({"name": name, "location": loc})
                    except OSError:
                        continue
                try:
                    winreg.CloseKey(key)
                except OSError:
                    pass
    except Exception:
        pass
    return results


def _find_testo_exes():
    """在常见安装目录+注册表AppPaths+开始菜单快捷方式中搜索 Testo/ComSoft 可执行文件（供一键启动）"""
    exes = []
    seen = set()
    if sys.platform != "win32":
        return exes

    def _add(p):
        try:
            p = os.path.normpath(p)
            pl = p.lower()
            if os.path.isfile(p) and pl.endswith(".exe") and pl not in seen:
                seen.add(pl)
                exes.append(p)
        except Exception:
            pass

    # 1) 注册表 App Paths（最可靠：安装即登记）
    try:
        import winreg
        for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            for wow in (winreg.KEY_WOW64_32KEY, winreg.KEY_WOW64_64KEY, 0):
                try:
                    k = winreg.OpenKey(root, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths",
                                       0, winreg.KEY_READ | wow)
                except OSError:
                    continue
                i = 0
                while True:
                    try:
                        sub = winreg.EnumKey(k, i)
                    except OSError:
                        break
                    i += 1
                    if "cc4" not in sub.lower() and "comsoft" not in sub.lower() and "testo" not in sub.lower():
                        continue
                    try:
                        sk = winreg.OpenKey(k, sub)
                        p = str(winreg.QueryValueEx(sk, "")[0])
                    except OSError:
                        continue
                    p = (p or "").strip().strip('"')
                    if p:
                        _add(p)
    except Exception:
        pass

    # 2) 已装软件 InstallLocation 深度扫描
    softwares = _find_testo_software()
    for sw in softwares:
        loc = sw.get("location") or ""
        if loc and os.path.isdir(loc):
            for root, dirs, files in os.walk(loc):
                if root.count(os.sep) - os.path.normpath(loc).count(os.sep) > 4:
                    dirs[:] = []
                    continue
                for f in files:
                    fl = f.lower()
                    if fl.endswith(".exe") and ("comsoft" in fl or "testo" in fl or fl.startswith("cc4")):
                        _add(os.path.join(root, f))

    # 3) 开始菜单快捷方式（读取 .lnk 内嵌路径）
    for base in ("ProgramData", "APPDATA"):
        b = os.environ.get(base)
        if not b:
            continue
        sm = os.path.join(b, "Microsoft", "Windows", "Start Menu")
        if not os.path.isdir(sm):
            continue
        for root, dirs, files in os.walk(sm):
            for f in files:
                if not f.lower().endswith(".lnk"):
                    continue
                fpath = os.path.join(root, f)
                if "comsoft" not in f.lower() and "testo" not in f.lower() and "cc4" not in f.lower():
                    continue
                try:
                    with open(fpath, "rb") as fh:
                        raw = fh.read()
                    try:
                        s = raw.decode("utf-16-le", "ignore")
                    except Exception:
                        s = ""
                    for tok in ("cc4.exe", "ComSoft.exe", "Comsoft.exe", "Testo.exe",
                                "cc4", "ComSoft", "Comsoft", "Testo"):
                        idx = s.lower().find(tok.lower())
                        if idx < 0:
                            continue
                        lo = s.rfind(":", 0, idx)
                        if lo < 0:
                            lo = max(0, idx - 60)
                        hi = s.find(chr(0), idx)
                        seg = s[lo:hi] if hi > lo else s[lo:]
                        seg = seg.replace(chr(0), "")
                        cand = seg.strip().strip('"')
                        if cand and cand.lower().endswith(".exe"):
                            _add(cand)
                except Exception:
                    continue

    # 4) 原有常见目录扫描
    roots = []
    for env in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432", "LOCALAPPDATA", "APPDATA"):
        base = os.environ.get(env)
        if not base:
            continue
        for sub in ("Testo", "testo", "ComSoft", "Comsoft"):
            roots.append(os.path.join(base, sub))
    for root_dir in roots:
        if not os.path.isdir(root_dir):
            continue
        for root, dirs, files in os.walk(root_dir):
            if root.count(os.sep) - root_dir.count(os.sep) > 3:
                dirs[:] = []
                continue
            for f in files:
                fl = f.lower()
                if fl.endswith(".exe") and ("comsoft" in fl or "testo" in fl or fl.startswith("cc4")):
                    _add(os.path.join(root, f))

    return exes


_CC4_PATH_FILE = os.path.join(APP_DATA_DIR, "cc4_path.txt")


def _saved_cc4_path():
    try:
        if os.path.isfile(_CC4_PATH_FILE):
            with open(_CC4_PATH_FILE, "r", encoding="utf-8") as f:
                p = f.read().strip().strip('"')
            if p and os.path.isfile(p):
                return p
    except Exception:
        pass
    return ""


def _save_cc4_path(p):
    try:
        os.makedirs(APP_DATA_DIR, exist_ok=True)
        with open(_CC4_PATH_FILE, "w", encoding="utf-8") as f:
            f.write((p or "").strip().strip('"'))
    except Exception:
        pass


@app.route("/api/comsoft/get-path", methods=["GET"])
def comsoft_get_path():
    return jsonify({"ok": True, "path": _saved_cc4_path()})


@app.route("/api/comsoft/set-path", methods=["POST"])
def comsoft_set_path():
    data = request.get_json(silent=True) or {}
    p = (data.get("exe") or "").strip().strip('"')
    if not p:
        return jsonify({"error": "路径为空"}), 400
    if not os.path.isfile(p):
        return jsonify({"error": "文件不存在: %s" % p}), 404
    _save_cc4_path(p)
    return jsonify({"ok": True, "saved": p})


@app.route("/api/comsoft/detect", methods=["GET"])
def comsoft_detect():
    softwares = _find_testo_software()
    exes = _find_testo_exes()
    # 合并去重：注册表结果 + 目录扫描结果
    known_paths = set()
    for sw in softwares:
        known_paths.add(sw.get("location", "").lower())
    for e in exes:
        # exe 已在某软件目录下则跳过
        if any(e.lower().startswith(p) for p in known_paths if p):
            continue
        softwares.append({"name": os.path.basename(e), "location": os.path.dirname(e), "exe": e})
    return jsonify({"platform": sys.platform, "softwares": softwares, "saved": _saved_cc4_path()})


@app.route("/api/comsoft/launch", methods=["POST"])
def comsoft_launch():
    """一键启动电脑上的 Testo / ComSoft 软件（仅 Windows）"""
    data = request.get_json(silent=True) or {}
    target = (data.get("exe") or "").strip().strip('"')
    if not target:
        # 1) 优先用户手动保存过的路径
        target = _saved_cc4_path()
    if not target:
        # 2) 增强自动扫描
        exes = _find_testo_exes()
        # cc4.exe 优先，其次 comsoft
        def _rank(p):
            pl = p.lower()
            if pl.endswith("cc4.exe"): return 0
            if "cc4" in pl: return 1
            if "comsoft" in pl: return 2
            return 3
        exes.sort(key=_rank)
        if exes:
            target = exes[0]
    if not target or not os.path.isfile(target):
        return jsonify({"error": "未找到 Testo/ComSoft 软件的可执行文件，请手动打开软件"}), 404
    try:
        subprocess.Popen([target], close_fds=True)
        return jsonify({"ok": True, "launched": target})
    except Exception as e:
        return jsonify({"error": f"启动失败: {e}"}), 500


def _recommended_export_dir():
    """推荐给用户在 ComSoft 里保存/导出数据的接力目录（自动创建）"""
    home = os.path.expanduser("~")
    candidates = [
        os.path.join(home, "Desktop", "testo_export"),
        os.path.join(home, "桌面", "testo_export"),
        os.path.join(home, "Documents", "testo_export"),
        os.path.join(home, "文档", "testo_export"),
        os.path.join(home, "testo_export"),
    ]
    for d in candidates:
        if os.path.isdir(os.path.dirname(d)):
            try:
                os.makedirs(d, exist_ok=True)
                return d
            except OSError:
                continue
    d = os.path.join(home, "testo_export")
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        pass
    return d


@app.route("/api/export-folder", methods=["GET"])
def export_folder_get():
    return jsonify({"path": _recommended_export_dir()})


@app.route("/api/export-folder/open", methods=["POST"])
def export_folder_open():
    d = _recommended_export_dir()
    try:
        if sys.platform == "win32":
            os.startfile(d)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", d])
        else:
            subprocess.Popen(["xdg-open", d])
        return jsonify({"ok": True, "path": d})
    except Exception as e:
        return jsonify({"error": f"打开文件夹失败: {e}", "path": d}), 500


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


# ─────────────────────────── v3.0.0 多设备批量插拔读取 ───────────────────────────
BATCH = {"id": None, "devices": [], "running": False}

# ─── 采集引擎（进程内驱动采集脚本，单 exe 内运行，自动导出 .vi2）─────────────────
COLLECTOR = {"started": 0, "log": []}


def _collector_start(folder=None, comsoft=None, watch=True):
    """在后台线程运行采集脚本(collector_driver)，驱动 ComSoft 自动导出 vi2。
    folder=输出目录(接力文件夹)；comsoft=cc4 路径。返回 (ok, 消息)。"""
    try:
        from collector import collector_driver as cd
    except Exception as e:
        COLLECTOR["log"].append("导入采集驱动失败: %s" % e)
        try:  # 开发模式直接路径
            import importlib.util
            _modp = os.path.join(BASE_DIR, "collector", "collector_driver.py")
            spec = importlib.util.spec_from_file_location("collector_driver", _modp)
            cd = importlib.util.module_from_spec(spec)
            sys.modules["collector_driver"] = cd
            spec.loader.exec_module(cd)
        except Exception as e2:
            COLLECTOR["log"].append("导入采集驱动失败(备选): %s" % e2)
            return False, "导入采集驱动失败"
    folder = folder or _ensure_desktop_raw_folder()
    comsoft = comsoft or _saved_cc4_path()
    try:
        ok, msg = cd.start(folder=folder, comsoft=comsoft, force=False)
        COLLECTOR["log"].append(msg)
        COLLECTOR["started"] = int(time.time())
        return ok, msg
    except Exception as e:
        COLLECTOR["log"].append("启动采集失败: %s" % e)
        return False, "启动采集失败: %s" % e


def _collector_stop():
    try:
        from collector import collector_driver as cd
        cd.stop()
        return True
    except Exception:
        try:
            import collector_driver as cd
            cd.stop()
            return True
        except Exception as e:
            return False


@app.route("/api/collector/status", methods=["GET"])
def collector_status():
    running = False
    logs = []
    try:
        from collector import collector_driver as cd
        st = cd.status()
        running = st.get("running", False)
        logs = st.get("log") or []
    except Exception:
        try:
            import collector_driver as cd
            st = cd.status()
            running = st.get("running", False)
            logs = st.get("log") or []
        except Exception:
            pass
    all_logs = list(logs) + list(COLLECTOR.get("log", [])[-10:])
    seen = set()
    merged = []
    for x in reversed(all_logs):
        if x not in seen:
            seen.add(x)
            merged.append(x)
        if len(merged) >= 20:
            break
    merged.reverse()
    return jsonify({"running": running, "log": merged})


@app.route("/api/collector/start", methods=["POST"])
def collector_start_api():
    data = request.get_json(silent=True) or {}
    folder = data.get("folder") or _ensure_desktop_raw_folder()
    comsoft = data.get("comsoft") or _saved_cc4_path()
    if not comsoft:
        return jsonify({"error": "尚未设置 ComSoft 路径，请先在设置中填写 cc4.exe 位置"}), 400
    ok, msg = _collector_start(folder, comsoft, watch=True)
    if not ok:
        return jsonify({"error": msg}), 500
    return jsonify({"ok": True, "msg": msg, "folder": folder})


@app.route("/api/collector/stop", methods=["POST"])
def collector_stop_api():
    _collector_stop()
    return jsonify({"ok": True})


def _batch_source(dev):
    if dev.get("_relay"):
        return "ComSoft全量存档(接力文件夹)"
    if dev.get("csv_files"):
        return "CSV"
    try:
        if any(f.lower().endswith(".vi2")
               for _, _, fs in os.walk(dev["path"]) for f in fs) if os.path.isdir(dev["path"]) else False:
            return "vi2存档(ComSoft导出)"
    except OSError:
        pass
    return "PDF报告(每分钟)" if dev.get("pdf_files") else "无"


def _batch_temp_val(temp):
    try:
        t = str(temp).replace(",", ".").replace("\u00b0C", "").replace("\u00b0F", "").strip()
        return float(t)
    except (ValueError, AttributeError):
        return None


def _batch_rec_time(r):
    d = r.get("date") or ""
    t = r.get("time") or ""
    if ":" in d:
        return d
    return (d + " " + t).strip()


@app.route("/api/batch/start", methods=["POST"])
def batch_start():
    import time as _t
    BATCH["id"] = str(int(_t.time() * 1000))
    BATCH["devices"] = []
    BATCH["running"] = True
    folder = _ensure_desktop_raw_folder()
    # 若已配置 ComSoft 且未在采集，自动启动采集引擎：用户只需插拔温度计
    collector_alive = False
    try:
        from collector import collector_driver as _cd
        collector_alive = bool(_cd.status().get("running"))
    except Exception:
        pass
    collector_started = False
    if not collector_alive and _saved_cc4_path():
        try:
            ok, _m = _collector_start(folder, _saved_cc4_path(), watch=True)
            collector_started = ok
        except Exception:
            collector_started = False
    return jsonify({"ok": True, "batch_id": BATCH["id"], "raw_folder": folder,
                    "collector_started": collector_started})


def _ensure_desktop_raw_folder():
    """在桌面创建「温度计原始数据-日期」文件夹，作为 ComSoft 接力导出目录并返回路径"""
    try:
        home = os.path.expanduser("~")
        desktop = None
        for cand in (os.path.join(home, "Desktop"), os.path.join(home, "桌面"),
                     os.path.join(home, "OneDrive", "桌面"), os.path.join(home, "OneDrive", "Desktop")):
            if os.path.isdir(cand):
                desktop = cand
                break
        if desktop is None:
            desktop = home
        folder = os.path.join(desktop, "温度计原始数据-%s" % datetime.now().strftime("%Y-%m-%d"))
        os.makedirs(folder, exist_ok=True)
        return folder
    except Exception:
        return ""


@app.route("/api/batch/detect", methods=["POST"])
def batch_detect():
    if not BATCH.get("running"):
        return jsonify({"error": "批次尚未开始，请先点击开始批量读取"}), 400
    try:
        devices = DeviceDetector.scan()
    except Exception as e:
        return jsonify({"error": f"检测失败: {e}"}), 500
    if not devices:
        return jsonify({"detected": False, "message": "未检测到设备，请插入下一台温度计"})
    # 优先: 已插入的 U 盘设备（非 relay）
    dev = next((d for d in devices if not d.get("_relay")), None)
    if dev is None:
        dev = devices[0]
    sn = dev.get("serial_number", dev["name"])
    # 若该 SN 在 ComSoft 接力文件夹存在全量 vi2/CSV 存档，自动升级为全量数据
    full = next((d for d in devices
                 if d.get("_relay") and str(d.get("serial_number", "")).strip() == str(sn).strip()), None)
    if full is not None and len(full.get("records") or []) > len(dev.get("records") or []):
        dev = full
        sn = dev.get("serial_number", dev["name"])
    recs = dev.get("records") or []
    temps = [_batch_temp_val(r.get("temperature")) for r in recs]
    temps = [t for t in temps if t is not None]
    seen = {d["sn"] for d in BATCH["devices"]}
    return jsonify({
        "detected": True,
        "already": sn in seen,
        "device": {
            "sn": sn,
            "name": dev["name"],
            "path": dev["path"],
            "record_count": len(recs),
            "source": _batch_source(dev),
            "start_time": _batch_rec_time(recs[0]) if recs else None,
            "end_time": _batch_rec_time(recs[-1]) if recs else None,
            "temp_min": round(min(temps), 2) if temps else None,
            "temp_max": round(max(temps), 2) if temps else None,
            "temp_avg": round(sum(temps) / len(temps), 2) if temps else None,
            "preview": recs[:5],
            "scan_log": dev.get("scan_log", []),
        },
    })


@app.route("/api/batch/save", methods=["POST"])
def batch_save():
    data = request.get_json(silent=True) or {}
    sn = (data.get("sn") or "").strip()
    try:
        devices = DeviceDetector.scan()
    except Exception as e:
        return jsonify({"error": f"重新读取设备失败: {e}"}), 500
    dev = None
    # 1) 优先精确匹配已插入设备；2) 再找该 SN 的全量 ComSoft 存档
    for d in devices:
        ds = d.get("serial_number", d["name"])
        if (ds == sn or (sn and sn in str(ds))) and not d.get("_relay"):
            dev = d
            break
    if dev is not None:
        full = next((d for d in devices
                     if d.get("_relay") and str(d.get("serial_number", "")).strip() == str(dev["serial_number"]).strip()),
                    None)
        if full is not None and len(full.get("records") or []) > len(dev.get("records") or []):
            dev = full
    if dev is None:
        # 3) 兜底：设备已拔，但接力文件夹有该 SN 的全量存档
        for d in devices:
            ds = d.get("serial_number", d["name"])
            if ds == sn or (sn and sn in str(ds)):
                dev = d
                break
    if not dev:
        return jsonify({"error": "设备已断开，请重新插入后检测"}), 400
    ds = dev.get("serial_number", dev["name"])
    if any(x["sn"] == ds for x in BATCH["devices"]):
        return jsonify({"ok": True, "duplicate": True, "device_count": len(BATCH["devices"])})
    BATCH["devices"].append({
        "sn": ds,
        "records": dev.get("records") or [],
        "source": _batch_source(dev),
        "unit": dev.get("unit", "°C"),
        "limit_min": dev.get("limit_min"),
        "limit_max": dev.get("limit_max"),
        "start_time": dev.get("start_time", ""),
    })
    return jsonify({
        "ok": True, "duplicate": False,
        "device_count": len(BATCH["devices"]),
        "record_count": len(dev.get("records") or []),
        "sn": ds,
    })


@app.route("/api/batch/autoscan", methods=["POST"])
def batch_autoscan():
    """一键全自动：扫描接力文件夹+已插设备，把新增测点(未在BATCH中)自动读入并去重。"""
    if not BATCH.get("running"):
        return jsonify({"ok": True, "added": [], "count": 0})
    try:
        devices = DeviceDetector.scan()
    except Exception as e:
        return jsonify({"error": "扫描失败: %s" % e}), 500
    best = {}
    for dev in devices:
        try:
            ds = dev.get("serial_number", dev["name"])
        except Exception:
            continue
        recs = dev.get("records") or []
        if not recs:
            continue
        if ds not in best or len(recs) > len(best[ds][0]):
            best[ds] = (recs, dev)
    seen = {d["sn"] for d in BATCH["devices"]}
    added = []
    for ds, (recs, dev) in best.items():
        if ds in seen:
            continue
        BATCH["devices"].append({
            "sn": ds,
            "records": recs,
            "source": _batch_source(dev),
            "unit": dev.get("unit", "°C"),
            "limit_min": dev.get("limit_min"),
            "limit_max": dev.get("limit_max"),
            "start_time": dev.get("start_time", ""),
        })
        seen.add(ds)
        added.append(ds)
        # 已读入的接力 vi2 归档到「已读/」子目录，避免下次批次重复读入历史积压
        try:
            if dev.get("_relay") and dev.get("path"):
                _src = dev["path"]
                if str(_src).lower().endswith(".vi2") and os.path.isfile(_src):
                    _dst_dir = os.path.join(os.path.dirname(_src), "已读")
                    os.makedirs(_dst_dir, exist_ok=True)
                    _dst = os.path.join(_dst_dir, os.path.basename(_src))
                    if _dst != _src and not os.path.exists(_dst):
                        os.rename(_src, _dst)
        except Exception:
            pass
    return jsonify({"ok": True, "added": added, "count": len(BATCH["devices"])})


@app.route("/api/batch/status", methods=["GET"])

def batch_status():
    return jsonify({
        "running": BATCH.get("running", False),
        "batch_id": BATCH.get("id"),
        "batch_ids": [d["sn"] for d in BATCH["devices"]],
        "devices": [
            {"sn": d["sn"], "record_count": len(d["records"]), "source": d["source"]}
            for d in BATCH["devices"]
        ],
    })


@app.route("/api/batch/export", methods=["POST"])
def batch_export():
    devices = [d for d in BATCH["devices"] if d["records"]]
    if not devices:
        return jsonify({"error": "批次中还没有已读取的设备数据，请先逐台检测并保存"}), 400
    try:
        path = _export_batch_excel(devices)
    except Exception as e:
        return jsonify({"error": f"导出失败: {e}"}), 500
    return jsonify({"ok": True, "file": os.path.basename(path), "path": path,
                    "device_count": len(devices)})


@app.route("/api/batch/clear", methods=["POST"])
def batch_clear():
    BATCH["running"] = False
    BATCH["devices"] = []
    _collector_stop()  # 结束采集，避免残留进程
    return jsonify({"ok": True})


@app.route("/api/batch/download/<filename>", methods=["GET"])
def batch_download(filename):
    safe = os.path.basename(filename)
    fp = os.path.join(EXPORT_DIR, safe)
    if not os.path.isfile(fp):
        return jsonify({"error": "文件不存在"}), 404
    return send_file(fp, as_attachment=True, download_name=safe)


def _export_batch_excel(devices):
    """导出：每个测点一个独立 xlsx(测点N) + 一个按时间对齐的汇总-演示-日期 xlsx，打包成 zip"""
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    import zipfile

    def _tval(r):
        return _batch_temp_val(r.get("temperature"))

    # 每个测点（vi2/设备）都保留，至少要有能读到温度的记录；不按 SN 去重，
    # 因为不同测点文件即使 SN 相同（如同一型号设备不同批次）也是独立测点，不能合并丢失。
    valid = []
    for d in devices:
        n = sum(1 for r in d.get("records") or [] if _tval(r) is not None)
        if n > 0:
            valid.append(d)
    if not valid:
        raise ValueError("没有可用数据")

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    today = datetime.now().strftime("%Y-%m-%d")
    os.makedirs(EXPORT_DIR, exist_ok=True)
    tmpdir = os.path.join(EXPORT_DIR, "tmp_%s" % stamp)
    os.makedirs(tmpdir, exist_ok=True)

    created = []  # (文件名, 路径)

    # ── 每个测点一个独立 xlsx：测点N.xlsx（头部信息块 + 时间/温度） ──
    for idx, d in enumerate(valid, 1):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "测点%d" % idx
        unit = str(d.get("unit") or "°C").strip() or "°C"
        sn = str(d.get("sn") or d.get("serial_number") or "").strip()
        # 记录时间列表
        times = []
        temps = []
        for r in d.get("records") or []:
            t = _tval(r)
            if t is None:
                continue
            times.append(_batch_rec_time(r))
            temps.append(t)
        # 统计
        tmin = round(min(temps), 2) if temps else ""
        tmax = round(max(temps), 2) if temps else ""
        tavg = round(sum(temps) / len(temps), 3) if temps else ""
        lm = d.get("limit_min")
        lx = d.get("limit_max")
        limit_str = ""
        if lm is not None and lx is not None:
            limit_str = "%s/%s" % (lm, lx)
        start_t = d.get("start_time") or (times[0] if times else "")
        end_t = times[-1] if times else ""
        exp_time = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        # 通道名+单位，如 "no name [°C]"；设备名称用 SN
        chan = "no name [%s]" % unit
        # ---- 头部信息块（参照样表） ----
        # 行1 设备名称 + 导出时间
        ws["A1"] = "设备名称: " + (sn or "未知")
        ws["E1"] = exp_time
        # 行2 起始时间 + 统计标签
        ws["A2"] = "起始时间: " + start_t.replace("-", "/")
        ws["C2"], ws["D2"], ws["E2"], ws["F2"] = "最小值", "最大值", "均值", "极限值"
        # 行3 结束时间 + 统计值
        ws["A3"] = "结束时间: " + (str(end_t).replace("-", "/") if end_t else "")
        ws["B3"] = chan
        ws["C3"], ws["D3"], ws["E3"], ws["F3"] = tmin, tmax, tavg, limit_str
        # 行4/5/6
        ws["A4"] = "测量通道: 1"
        ws["A5"] = "测量值: %d" % len(temps)
        ws["A6"] = "SN %s" % (sn or "未知")
        # 样式：头部标签加粗，统计表头加粗蓝底
        for rr, cnames in ((2, ("C2", "D2", "E2", "F2")),):
            for c in cnames:
                ws[c].font = Font(bold=True, color="FFFFFF")
                ws[c].fill = PatternFill("solid", fgColor="4472C4")
        for c in ("A1", "A2", "A3", "A4", "A5", "A6", "B3"):
            ws[c].font = Font(bold=True)
        # 空行第7行
        # 行8 数据表头
        hdr_row = 8
        ws.cell(hdr_row, 1, "id")
        ws.cell(hdr_row, 2, "日期/时间")
        ws.cell(hdr_row, 3, chan)
        for c in range(1, 4):
            cc = ws.cell(hdr_row, c)
            cc.font = Font(bold=True, color="FFFFFF")
            cc.fill = PatternFill("solid", fgColor="4472C4")
        # 数据行
        for i, (tm, tv) in enumerate(zip(times, temps), 1):
            ws.cell(hdr_row + i, 1, i)
            ws.cell(hdr_row + i, 2, tm)
            ws.cell(hdr_row + i, 3, tv)
        ws.column_dimensions["A"].width = 16
        ws.column_dimensions["B"].width = 22
        ws.column_dimensions["C"].width = 16
        ws.column_dimensions["D"].width = 12
        ws.column_dimensions["E"].width = 12
        ws.column_dimensions["F"].width = 12
        fname = "测点%d.xlsx" % idx
        p = os.path.join(tmpdir, fname)
        wb.save(p)
        created.append((fname, p))

    # ── 汇总-演示-日期.xlsx：按时间对齐，缺数据留空 ──
    wb = openpyxl.Workbook()
    summary = wb.active
    summary.title = "汇总数据"
    summary.append(["温度计号"] + [str(d.get("sn") or d.get("serial_number") or "") for d in valid])
    summary.append(["日期"] + ["测点%d" % i for i in range(1, len(valid) + 1)])
    for col in range(1, len(valid) + 2):
        c1 = summary.cell(1, col)
        c2 = summary.cell(2, col)
        c1.font = Font(bold=True, color="FFFFFF")
        c2.font = Font(bold=True, color="FFFFFF")
        c1.fill = PatternFill("solid", fgColor="4472C4")
        c2.fill = PatternFill("solid", fgColor="4472C4")
    # 各测点：分钟key -> 温度（同批测点起始秒不同，但采样为整分钟，按分钟对齐）
    maps = []
    for d in valid:
        m = {}
        for r in d.get("records") or []:
            t = _tval(r)
            if t is None:
                continue
            key = _batch_rec_time(r)[:16]  # 分钟级对齐，忽略起始秒差异
            m[key] = t
        maps.append(m)
    # 时间轴 = 所有测点分钟的「合集」：各测点起止时间可能有偏差，取全部测点出现的
    # 分钟去重并升序作为首列，每个测点按该分钟匹配填入温度，无数据则留空，不改原始数据
    all_times = set()
    for m in maps:
        all_times.update(m.keys())
    for t in sorted(all_times):
        row = [t]
        for m in maps:
            row.append(m.get(t, ""))
        summary.append(row)
    summary.column_dimensions["A"].width = 34
    for col in range(2, len(valid) + 2):
        summary.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 12
    sum_fname = "汇总-演示-%s.xlsx" % today
    sum_path = os.path.join(tmpdir, sum_fname)
    wb.save(sum_path)
    created.append((sum_fname, sum_path))

    # ── 打包 zip ──
    zip_path = os.path.join(EXPORT_DIR, "Testo184_汇总_%s.zip" % stamp)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, p in created:
            zf.write(p, fname)
    # 清理临时文件
    for _, p in created:
        try:
            os.remove(p)
        except OSError:
            pass
    try:
        os.rmdir(tmpdir)
    except OSError:
        pass
    return zip_path


# ─── v3.1.0 批量上传 .vi2 → 一键导出 Excel（独立于 session，直接全量解析） ───

# 内存暂存：一次会话内上传的 vi2 设备（sn -> 数据）
VI2_BATCH = {"devices": []}  # [{"sn","records","source","file"}]


@app.route("/api/vi2/upload", methods=["POST"])
def vi2_batch_upload():
    """一次上传多个 .vi2，逐个全量解析，返回设备列表；同名 SN 不重复加入。"""
    files = request.files.getlist("file")
    if not files:
        return jsonify({"error": "未收到文件"}), 400
    tmp_dir = os.path.join(APP_DATA_DIR, "tmp")
    os.makedirs(tmp_dir, exist_ok=True)
    # 按“源文件名”去重，而不是 SN：同一型号设备不同批次/文件即使 SN 相同也是独立测点
    dexists = {d.get("file") for d in VI2_BATCH["devices"]}
    added, errors = [], []
    for file in files:
        fname = file.filename or ""
        if not fname.lower().endswith(".vi2"):
            errors.append({"file": fname, "error": "仅支持 .vi2 文件"})
            continue
        tmp_path = os.path.join(tmp_dir, f"{uuid.uuid4().hex}.vi2")
        try:
            file.save(tmp_path)
            data = parse_vi2(tmp_path)
            sn = str(data.get("serial_number") or "未知").strip()
            recs = data.get("records") or []
            if not recs:
                errors.append({"file": fname, "error": "未解析到数据"})
                continue
            if fname in dexists:
                errors.append({"file": fname, "error": f"文件 {fname} 已上传，跳过重复"})
                continue
            VI2_BATCH["devices"].append({
                "sn": sn, "records": recs, "file": fname,
                "source": "vi2全量(批量上传)",
                "unit": data.get("unit", "°C"),
                "limit_min": data.get("limit_min"),
                "limit_max": data.get("limit_max"),
                "start_time": data.get("start_time", ""),
            })
            dexists.add(fname)
            added.append({
                "file": fname, "sn": sn, "record_count": len(recs),
                "unit": data.get("unit", "°C"),
            })
        except Exception as e:
            errors.append({"file": fname, "error": str(e)})
        finally:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
    return jsonify({
        "ok": bool(added),
        "added": added,
        "errors": errors,
        "device_count": len(VI2_BATCH["devices"]),
    })


@app.route("/api/vi2/list", methods=["GET"])
def vi2_batch_list():
    return jsonify({"devices": [
        {"sn": d["sn"], "record_count": len(d["records"]), "source": d["source"], "file": d["file"]}
        for d in VI2_BATCH["devices"]
    ]})


@app.route("/api/vi2/clear", methods=["POST"])
def vi2_batch_clear():
    VI2_BATCH["devices"] = []
    return jsonify({"ok": True})


@app.route("/api/vi2/export", methods=["POST"])
def vi2_batch_export():
    devices = [d for d in VI2_BATCH["devices"] if d.get("records")]
    if not devices:
        return jsonify({"error": "还没有已上传的设备，请先上传 .vi2 文件"}), 400
    try:
        path = _export_batch_excel(devices)
    except Exception as e:
        return jsonify({"error": f"导出失败: {e}"}), 500
    return jsonify({"ok": True, "file": os.path.basename(path), "path": path,
                    "device_count": len(devices)})


# ─── v3.2.0 统一数据池：两种读取方式(设备读取/vi2批量导入)合并，一键导出 ───

def _unified_devices():
    """合并 设备读取(BATCH) + vi2批量导入(VI2_BATCH)。

    注意：不按 SN 去重/合并——同一型号设备可能被读取成多个测点（不同文件/批次），
    SN 相同不代表是同一个测点。这里按“文件唯一键”合并，避免同文件被不同来源重复加入，
    但保留所有不同文件（不同测点）的设备。
    """
    merged = {}
    def _key(d):
        # 优先用具体文件路径/文件名；都没有才退回 SN
        return (d.get("path") or d.get("file") or d.get("name") or
                ("SN:" + str(d.get("sn") or "")))
    def _recs(d):
        return d.get("records") or []
    for d in list(BATCH.get("devices", [])) + list(VI2_BATCH.get("devices", [])):
        key = _key(d)
        if not key:
            continue
        cur = merged.get(key)
        if cur is None or len(_recs(d)) > len(_recs(cur)):
            merged[key] = dict(d)
    return list(merged.values())


@app.route("/api/all/list", methods=["GET"])
def all_devices_list():
    return jsonify({"devices": [
        {"sn": d.get("sn"), "record_count": len(d.get("records") or []),
         "source": d.get("source", ""), "file": d.get("file", "")}
        for d in _unified_devices()
    ]})


@app.route("/api/all/export", methods=["POST"])
def all_devices_export():
    devices = [d for d in _unified_devices() if d.get("records")]
    if not devices:
        return jsonify({"error": "还没有数据。请在「读取设备」或「批量导入vi2」中先读取/导入数据。"}), 400
    try:
        path = _export_batch_excel(devices)
    except Exception as e:
        return jsonify({"error": f"导出失败: {e}"}), 500
    fname = os.path.basename(path)
    total = sum(len(d.get("records") or []) for d in devices)
    return jsonify({"ok": True, "file": fname, "path": path,
                    "device_count": len(devices), "record_count": total,
                    "count": len(devices), "records": total,
                    "filename": fname,
                    "url": "/api/batch/download/" + fname})

@app.route("/api/all/clear", methods=["POST"])
def all_devices_clear():
    BATCH["devices"] = []
    BATCH["running"] = False
    VI2_BATCH["devices"] = []
    return jsonify({"ok": True})


def _find_desktop():
    """返回桌面路径（兼容 OneDrive/中文桌面），找不到返回 None"""
    try:
        home = os.path.expanduser("~")
        for cand in (os.path.join(home, "Desktop"), os.path.join(home, "桌面"),
                     os.path.join(home, "OneDrive", "Desktop"), os.path.join(home, "OneDrive", "桌面")):
            if os.path.isdir(cand):
                return cand
        return home
    except Exception:
        return None


def _ensure_desktop_shortcut():
    """首次运行在桌面创建「Testo184 数据汇总工具.lnk」快捷方式指向本程序。仅 Windows + 打包后生效。"""
    if platform.system() != "Windows":
        return
    exe = os.path.abspath(sys.executable if getattr(sys, "frozen", False) else "")
    if not exe or not exe.lower().endswith(".exe"):
        return  # 源码运行不需要快捷方式
    desktop = _find_desktop()
    if not desktop:
        return
    lnk = os.path.join(desktop, "Testo184 数据汇总工具.lnk")
    if os.path.exists(lnk):
        return  # 已存在不重复创建
    # 用 PowerShell WScript.Shell 创建 .lnk（避免依赖第三方库）
    try:
        ps = (
            "$s=(New-Object -ComObject WScript.Shell).CreateShortcut('" + lnk + "');"
            "$s.TargetPath='" + exe + "';"
            "$s.WorkingDirectory='" + os.path.dirname(exe) + "';"
            "$s.IconLocation='" + exe + ",0';"
            "$s.Save()"
        )
        import base64
        enc = base64.b64encode(ps.encode("utf-16-le")).decode("ascii")
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                        "-EncodedCommand", enc], capture_output=True, timeout=30)
    except Exception:
        pass


def _ensure_autostart():
    """开机自启：写 HKCU Run 键。仅 Windows；Linux/macOS 跳过。"""
    if platform.system() != "Windows":
        return
    exe = os.path.abspath(sys.executable if getattr(sys, "frozen", False) else "")
    if not exe or not exe.lower().endswith(".exe"):
        return
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Windows\CurrentVersion\Run",
                             0, winreg.KEY_SET_VALUE)
        # 用参数 --service：开机启动时不重复打开浏览器
        winreg.SetValueEx(key, "Testo184Tool", 0, winreg.REG_SZ,
                          '"%s" --service' % exe)
        winreg.CloseKey(key)
    except Exception:
        pass


def _service_running(port):
    """探测端口是否已有服务在跑（避免重复启动/端口冲突）。"""
    try:
        import socket
        s = socket.create_connection(("127.0.0.1", int(port)), timeout=0.6)
        s.close()
        return True
    except OSError:
        return False


if __name__ == "__main__":
    init_db()
    is_service = "--service" in sys.argv  # 开机自启后台模式
    port = int(os.environ.get("TESTO_PORT", 8000))

    if not is_service:
        # 普通双击：自动建桌面快捷方式 + 注册开机自启
        _ensure_desktop_shortcut()
        _ensure_autostart()
        print()
        print("╔══════════════════════════════════════════════════╗")
        print("║   Testo 184 温度计数据读取汇总工具              ║")
        print("╠══════════════════════════════════════════════════╣")
        print("║                                                  ║")
        print("║   请在浏览器中打开:                              ║")
        print("║   ➜  http://localhost:%d                       ║" % port)
        print("║                                                  ║")
        print("║   已在本机桌面创建快捷方式，并设为开机自启       ║")
        print("║   之后开机自动运行，直接打开网页即可使用         ║")
        print("║                                                  ║")
        print("║   按 Ctrl+C 停止本次服务                         ║")
        print("╚══════════════════════════════════════════════════╝")
        print()
        # 若服务已在运行（开机自启已起）：只开浏览器，不重复启动服务
        if _service_running(port):
            def _open_only():
                import time
                time.sleep(0.5)
                import webbrowser
                try:
                    webbrowser.open("http://localhost:%d" % port)
                except Exception:
                    pass
            threading.Thread(target=_open_only, daemon=True).start()
            print("  服务已在后台运行，已为你打开浏览器。本窗口可以直接关闭。")
            import time as _t
            _t.sleep(2)
            sys.exit(0)
    else:
        # 开机自启后台模式：静默启动，不打印、不重复开浏览器
        if _service_running(port):
            sys.exit(0)

    # 自动在默认浏览器中打开（服务启动后延迟打开，避免端口未就绪）
    def _open_browser():
        import time
        time.sleep(1.0)
        import webbrowser
        try:
            webbrowser.open("http://localhost:%d" % port)
        except Exception:
            pass

    threading.Thread(target=_open_browser, daemon=True).start()
    app.run(host="0.0.0.0", port=port, debug=False)

