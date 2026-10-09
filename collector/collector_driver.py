# -*- coding: utf-8 -*-
"""采集驱动：在主程序进程内运行采集脚本（Windows 专用），封装线程/日志/停止。
采集脚本(RPA)作为资源打进主 exe，本模块在主程序后台线程中运行它，导出的 .vi2
写入接力文件夹，主程序同时监控该文件夹自动读入汇总。"""

import os
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

COLLECTOR_IMPORT = None      # 缓存的采集模块
COLLECTOR_THREAD = None      # 运行线程
COLLECTOR_STOP = None        # 停止事件
COLLECTOR_LOGS = []          # 最近日志
_COLLECTOR_LOCK = threading.Lock()


def _locate_collector_script():
    """定位采集脚本路径：打包后取 _MEIPASS，开发模式取项目 collector/"""
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            # 单 exe：脚本打进 _MEIPASS/collector/
            cand = os.path.join(meipass, "collector", "testo184_vi2_export.py")
            if os.path.isfile(cand):
                return cand
        # 兜底：exe 同目录
        exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        cand = os.path.join(exe_dir, "testo184_vi2_export.py")
        if os.path.isfile(cand):
            return cand
        return None
    # 开发模式
    base = os.path.dirname(os.path.abspath(__file__))
    cand = os.path.join(base, "testo184_vi2_export.py")
    return cand if os.path.isfile(cand) else None


def _load_collector():
    """import 采集模块（Windows 专属，延迟加载、容错）。"""
    global COLLECTOR_IMPORT
    if COLLECTOR_IMPORT is not None:
        return COLLECTOR_IMPORT
    script = _locate_collector_script()
    if not script:
        return None
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location("testo184_vi2_export", script)
        mod = importlib.util.module_from_spec(spec)
        # 预注入 sys.modules 以便其内部相对引用
        sys.modules["testo184_vi2_export"] = mod
        spec.loader.exec_module(mod)
        COLLECTOR_IMPORT = mod
        return mod
    except Exception as e:
        _log("加载采集模块失败: %s" % e)
        return None


def _log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    with _COLLECTOR_LOCK:
        COLLECTOR_LOGS.append("[%s] %s" % (ts, msg))
        if len(COLLECTOR_LOGS) > 200:
            del COLLECTOR_LOGS[:-100]


def _worker(cfg, base, force):
    mod = _load_collector()
    if mod is None:
        _log("未找到采集模块，无法自动采集")
        COLLECTOR_THREAD._result = "no_module"
        return
    # 采集脚本内部用 print 输出，主程序 noconsole 无窗口 → 重定向到日志
    _orig_print = print
    def _captured_print(*a, **k):
        text = " ".join(str(x) for x in a)
        _log(text)
    import builtins
    builtins.print = _captured_print
    try:
        # 让采集脚本的 stop 感知外部停止
        mod.stop_requested = COLLECTOR_STOP
        rc = mod.run_watch(Path(base), cfg, force)
        COLLECTOR_THREAD._result = "rc_%s" % rc
        _log("采集结束 (rc=%s)" % rc)
    except Exception as e:
        _log("采集异常: %s" % e)
        COLLECTOR_THREAD._result = "error"
    finally:
        builtins.print = _orig_print
        _log("采集进程已退出")


def start(folder=None, comsoft=None, force=False):
    """启动采集（线程运行）。folder=接力目录；comsoft=cc4 路径。返回 (ok,msg)。"""
    global COLLECTOR_THREAD, COLLECTOR_STOP
    if COLLECTOR_THREAD is not None and COLLECTOR_THREAD.is_alive():
        return True, "采集已在运行"
    mod = _load_collector()
    if mod is None:
        return False, "未找到采集模块"
    cfg = dict(getattr(mod, "DEFAULT_CONFIG", {}))
    if comsoft:
        cfg["comsoft_exe"] = os.path.expandvars(comsoft)
    COLLECTOR_STOP = threading.Event()
    COLLECTOR_THREAD = threading.Thread(target=_worker,
                                        args=(cfg, Path(folder or "."), force),
                                        daemon=True)
    COLLECTOR_THREAD.daemon = True
    COLLECTOR_THREAD.start()
    _log("采集已启动")
    return True, "采集已启动"


def stop():
    global COLLECTOR_THREAD, COLLECTOR_STOP
    if COLLECTOR_STOP:
        COLLECTOR_STOP.set()
    if COLLECTOR_THREAD and COLLECTOR_THREAD.is_alive():
        COLLECTOR_THREAD.join(timeout=3)
    COLLECTOR_THREAD = None
    COLLECTOR_STOP = None
    _log("已请求停止采集")
    return True


def status():
    running = bool(COLLECTOR_THREAD and COLLECTOR_THREAD.is_alive())
    return {"running": running, "log": COLLECTOR_LOGS[-30:]}
