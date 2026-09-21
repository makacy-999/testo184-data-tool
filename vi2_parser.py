# -*- coding: utf-8 -*-
"""Testo .vi2 数据文件解析器

.vi2 是 Testo ComSoft 专有的存档格式，实为 Microsoft Compound File (OLE2/CFBF)。
通过分析真实样本破解出的结构：

关键 streams:
  - <sn>/t18b             : 明文，DeviceType / SerialNumber
  - <sn>/data/values      : 测量数据，每条记录 = [4字节 uint32 时间码][4字节 float32 温度]
  - <sn>/data/schema      : 维度信息(通道数等)
  - <sn>/summary          : 起始/结束时间码 + 记录数
  - <sn>/channels/1/<meta>: 通道元数据(通道名/单位)
  - audittrail            : 设备操作日志(XML)，含采样间隔、单位等
"""

import olefile
import struct
from datetime import datetime, timedelta


def _extract_sn(ole):
    """从 t18b 提取 serial number 与 device type"""
    for path in ole.listdir():
        if path and path[-1].lower() == "t18b":
            try:
                text = ole.openstream(path).read().decode("utf-8", "replace")
                sn = ""
                dtype = ""
                for line in text.splitlines():
                    line = line.strip()
                    if line.lower().startswith("serialnumber"):
                        parts = line.replace("\t", " ").split(" ", 1)
                        if len(parts) > 1:
                            sn = parts[1].strip()
                    elif line.lower().startswith("devicetype"):
                        parts = line.replace("\t", " ").split(" ", 1)
                        if len(parts) > 1:
                            dtype = parts[1].strip()
                return sn, dtype
            except Exception:
                return "", ""
    return "", ""


def _find_data_dir(ole):
    """定位包含 data/values 的目录(形如 '29328')"""
    for path in ole.listdir():
        if len(path) >= 3 and path[1].lower() == "data" and path[2].lower() == "values":
            return path[0]
    return None


def _find_channels_unit(ole, data_dir):
    """尝试从通道元数据提取单位(如 °C)"""
    # 优先看 audittrail 里的 'C:1 °C' 模式
    if ole.exists("audittrail"):
        try:
            text = ole.openstream("audittrail").read().decode("utf-8", "replace")
            import re
            m = re.search(r"C:\d+\s*([^\s\"'>]+)", text)
            if m:
                return m.group(1)
        except Exception:
            pass
    return "°C"


def _parse_values(ole, data_dir):
    """解析 data/values：每 8 字节 = [uint32 时间码][float32 温度]"""
    values_path = [data_dir, "data", "values"]
    if not ole.exists(values_path):
        return []
    data = ole.openstream(values_path).read()
    n = len(data) // 8
    records = []
    for i in range(n):
        chunk = data[i * 8:(i + 1) * 8]
        if len(chunk) < 8:
            break
        t_code = struct.unpack("<I", chunk[0:4])[0]
        temp = struct.unpack("<f", chunk[4:8])[0]
        records.append({"t_code": t_code, "temperature": temp})
    return records


def parse_vi2(file_path, sample_minutes=None, start_time=None):
    """解析 .vi2 文件。

    返回 dict:
      {
        "serial_number": str,
        "device_type": str,
        "unit": str,
        "record_count": int,
        "sample_minutes": float,     # 采样间隔(分钟)，
        "start_time": str,           # 起始时间(YYYY-MM-DD HH:MM:SS)
        "records": [{"date": str, "time": str, "temperature": float, "t_code": int}, ...]
      }
    """
    ole = olefile.OleFileIO(file_path)
    try:
        sn, dtype = _extract_sn(ole)
        data_dir = _find_data_dir(ole)
        unit = _find_channels_unit(ole, data_dir)
        records = _parse_values(ole, data_dir) if data_dir else []

        # 采样间隔：未指定时从 audittrail 推断；Testo 184 常见 2/5/8 分钟
        if not sample_minutes:
            sample_minutes = _infer_sample_minutes(ole) or 2.0

        # 时间基站：可指定；否则用 audittrail 的最后操作时间倒推
        if not start_time:
            end_hint = _last_audit_time(ole)
            if end_hint and records:
                total_min = (records[-1]["t_code"] - records[0]["t_code"]) * (sample_minutes / 64.0)
                start_time = end_hint - timedelta(minutes=total_min)
            else:
                start_time = datetime.now().replace(second=0, microsecond=0)
        elif isinstance(start_time, str):
            start_time = datetime.strptime(start_time, "%Y-%m-%d %H:%M:%S")

        # 生成带时间的记录
        per_unit_min = sample_minutes / 64.0  # 每个时间码单位对应分钟数
        out = []
        for idx, r in enumerate(records):
            delta_min = (r["t_code"] - records[0]["t_code"]) * per_unit_min
            dt = start_time + timedelta(minutes=delta_min)
            out.append({
                "date": dt.strftime("%Y-%m-%d"),
                "time": dt.strftime("%H:%M:%S"),
                "temperature": round(r["temperature"], 2),
                "t_code": r["t_code"],
            })

        return {
            "serial_number": sn or "未知",
            "device_type": dtype or "",
            "unit": unit,
            "record_count": len(out),
            "sample_minutes": sample_minutes,
            "start_time": start_time.strftime("%Y-%m-%d %H:%M:%S"),
            "records": out,
        }
    finally:
        ole.close()


def _infer_sample_minutes(ole):
    """从 audittrail 推断采样间隔(分钟)，猜不到返回 None"""
    if not ole.exists("audittrail"):
        return None
    try:
        text = ole.openstream("audittrail").read().decode("utf-8", "replace")
        import re
        # 形如 '8.00   C:1 °C' 或 '2.00 C:1 °C'
        m = re.search(r"(\d+\.?\d*)\s+C:\d+", text)
        if m:
            val = float(m.group(1))
            if 0.1 <= val <= 1440:
                return val
    except Exception:
        pass
    return None


def _last_audit_time(ole):
    """从 audittrail 提取最后操作时间（作为记录结束时间的参考）"""
    if not ole.exists("audittrail"):
        return None
    try:
        text = ole.openstream("audittrail").read().decode("utf-8", "replace")
        import re
        times = re.findall(r"Date='([\dT:\.\-]+)'", text)
        for tstr in reversed(times):
            try:
                return datetime.strptime(tstr, "%Y-%m-%dT%H:%M:%S")
            except ValueError:
                continue
    except Exception:
        pass
    return None