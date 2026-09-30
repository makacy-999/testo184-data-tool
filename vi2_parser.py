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

        # ── 采样间隔：从 values 时间码的主流间隔推断，绝不用 audittrail 极限值 ──
        # 每 64 个时间码 = 1 分钟(固定分辨率)；主流相邻时间码差 ÷ 64 = 采样间隔分钟。
        # 例如主流间隔 64 → 1 分钟/条，128 → 2 分钟/条。
        # 注意：values 里偶有 320/33088 等大跳变，是设备采样中断的计数缺口，不是真实记录间隔，
        # 推导采样间隔时只取“主流间隔”(出现次数最多的差值)。
        if not sample_minutes:
            sample_minutes = _infer_sample_minutes_from_values(records) or 1.0

        # ── 测量起始时间：优先 alrtobject(测量开始的真实绝对时间) ──
        # alrtobject 的 c0 即测量开始时刻，是 ComSoft 导出 Excel 所用的权威锚点，
        # 比 audittrail(导出操作时刻)可靠得多。仅当无报警/无 alrtobject 时才回退。
        if not start_time:
            alrt_start, alrt_end = _extract_alrt_range(ole)
            if alrt_start:
                try:
                    start_time = datetime.strptime(
                        alrt_start[:19], "%Y-%m-%dT%H:%M:%S")
                except ValueError:
                    try:
                        start_time = datetime.strptime(
                            alrt_start[:19], "%Y-%m-%d %H:%M:%S")
                    except ValueError:
                        start_time = None
            if not start_time:
                # 回退：用 audittrail 最后操作时间 - 记录总跨度(仅当无 alrtobject)
                end_hint = _last_audit_time(ole)
                if end_hint and records:
                    start_time = end_hint - timedelta(minutes=(len(records) - 1) * sample_minutes)
                else:
                    start_time = datetime.now().replace(second=0, microsecond=0)
        elif isinstance(start_time, str):
            start_time = datetime.strptime(start_time, "%Y-%m-%d %H:%M:%S")

        # ── 生成带时间的记录：按记录条索引等间隔累加，而非按时间码差 ──
        # 记录采样是均匀的(每 sample_minutes 分钟一条)，故时间 = 起始时间 + 索引×采样间隔。
        # (若按 t_code 差累加，320/33088 的中断缺口会被错误计入真实经过时间，导致时间错乱。)
        out = []
        for idx, r in enumerate(records):
            dt = start_time + timedelta(minutes=idx * sample_minutes)
            out.append({
                "date": dt.strftime("%Y-%m-%d"),
                "time": dt.strftime("%H:%M:%S"),
                "temperature": round(r["temperature"], 2),
                "t_code": r["t_code"],
            })

        limit_min, limit_max = _extract_limits(ole)
        return {
            "serial_number": sn or "未知",
            "device_type": dtype or "",
            "unit": unit,
            "record_count": len(out),
            "sample_minutes": sample_minutes,
            "start_time": start_time.strftime("%Y-%m-%d %H:%M:%S"),
            "limit_min": limit_min,
            "limit_max": limit_max,
            "records": out,
        }

    finally:
        ole.close()

def _extract_alrt_range(ole):
    """从 alrtobject/rcrds 提取测量起止绝对时间(权威锚点)。

    alrtobject 是设备触发报警的时段记录，其 c0=测量开始时间、c1=测量结束时间，
    是 vi2 内最可靠的测量起止时间戳（ComSoft 导出 Excel 即以此为准）。
    返回 (start_str, end_str) 或 (None, None)。
    """
    try:
        for entry in ole.listdir():
            if 'alrtobject' in entry and entry[-1] == 'rcrds':
                text = ole.openstream(entry).read().decode("utf-8", "ignore")
                import re as _re
                m = _re.search(r"c0='([^']+)'[^>]*c1='([^']+)'", text)
                if m:
                    return m.group(1), m.group(2)
                m0 = _re.search(r"c0='([^']+)'", text)
                m1 = _re.search(r"c1='([^']+)'", text)
                if m0 and m1:
                    return m0.group(1), m1.group(2)
    except Exception:
        pass
    return None, None

def _extract_limits(ole):
    """从 audittrail 提取极限值(下限/上限)。Action 5768=下限, 5769=上限。找不到返回 (None,None)"""
    if not ole.exists("audittrail"):
        return None, None
    try:
        text = ole.openstream("audittrail").read().decode("utf-8", "ignore")
    except Exception:
        return None, None
    import re
    limit_min = limit_max = None
    # 匹配 Comments='25.00   C:1 °C' Action='5769' 或 Action 在前后
    rows = re.findall(r"Comments='([^']*)'[^>]*Action='(\d+)'", text)
    for comment, action in rows:
        m = re.search(r"([-+]?\d+(?:\.\d+)?)\s*C\s*:\s*\d+\s*[^\s']*", comment)
        if not m:
            m = re.search(r"([-+]?\d+(?:\.\d+)?)\s*°?\s*C", comment)
        if not m:
            continue
        try:
            val = float(m.group(1))
        except ValueError:
            continue
        if action == "5768" and limit_min is None:
            limit_min = val
        elif action == "5769" and limit_max is None:
            limit_max = val
    return limit_min, limit_max


# parse_vi2 的 finally 收尾在下文恢复


def _infer_sample_minutes_from_values(records):
    """从 values 时间码的主流相邻间隔推断采样间隔(分钟)。

    每 64 个时间码 = 1 分钟(固定分辨率)。取相邻 t_code 差中出现次数最多的值为主流间隔，
    采样间隔 = 主流间隔 / 64。中断缺口(320/33088 等)不计入主流，故不影响判断。
    """
    if not records:
        return None
    from collections import Counter
    gaps = Counter()
    for i in range(1, len(records)):
        g = records[i]["t_code"] - records[i - 1]["t_code"]
        if g > 0:
            gaps[g] += 1
    if not gaps:
        return None
    main_gap = gaps.most_common(1)[0][0]
    sample_min = main_gap / 64.0
    if 0.1 <= sample_min <= 1440:
        return sample_min
    return None


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