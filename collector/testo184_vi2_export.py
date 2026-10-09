#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
Testo 184 USB 温度计 · .vi2 数据文件自动导出（需 Testo Comfort Software）
========================================================================

背景
----
设备 U 盘里只有 PDF 报告；Comsoft 存档格式 .vi2（含原始测量数据）只能由
Testo Comfort Software 生成。本脚本通过 UI 自动化驱动 Comsoft 完成：

    双击存档树中设备的 "testo 184 measurement report.pdf"
    → 文件菜单"另存为" → 填入 <设备号>.vi2 → 保存 → 校验 → 关闭标签

关键机制（实测 2026-10-08）
--------------------------
* **绝不使用 pywinauto 的 uia 后端**：在本机环境中，
  `Desktop(backend="uia").windows()` 与 `element.descendants()` 会因
  UIA 提供者无响应而**永久死锁**（Ctrl+C 也无法中断）。这是本工具此前
  "按 4 后卡死、鼠标键盘不能用" 的根因。
* 改为 **comtypes 直接调用原生 UIA**，并用 `RawViewWalker` 逐层遍历。
* **找主窗口用 win32 EnumWindows（~11ms），不用 UIA root.FindFirst**：
  后者在本机会永久死锁。拿到 hwnd 后用 ElementFromHandle 换 UIA 元素。
* 存档树 = SysTreeView32，节点 `testo184-2013 <设备号>: <设备号>`，
  其下子节点 `testo 184 measurement report.pdf`。
* 打开文档：对 report 子节点的**屏幕坐标双击，只双击一遍**（Invoke 无效）。
* 菜单：点击"文件"按钮（物理点击）→ 点"另存为(A)..."菜单项
  （菜单项坐标在同一窗口会话内稳定 → 缓存复用，免去每轮 UIA 遍历）。
* 保存对话框"保存为"为 MFC 标准对话框：
    文件名 Edit = control_id 1152，保存 Button = control_id 1，取消 = 2。
  Edit 只接受 AttachThreadInput 后的 WM_SETTEXT，写入后必须回读校验。
* 设备号校验：.vi2 原始字节内含 `SerialNumber\t<设备号>`，直接字节搜索。
* **关闭文档 = 点文档右上角 ×**（给文档视图发 WM_CLOSE）。
  **绝不向主窗口发 SC_CLOSE** —— 没有文档时那会直接关掉整个 Comsoft。

性能（2 台设备端到端）
--------------------
优化前 24.4s → 优化后 **6.5~7s**。主要手法：
  * 全树 UIA 遍历 → win32 EnumChildWindows/EnumWindows（快 100~900 倍）
  * 固定 sleep → 高频轮询（0.01~0.02s）
  * 双击两遍 + 重试 → 单次双击
  * 关闭按钮每轮重定位（0.9s/轮）→ 直接关文档窗口（~45ms/个）
  * 按钮/菜单坐标缓存复用

要求
----
* Windows + Testo Comfort Software（中文界面）+ Python 3.7+
* 依赖：comtypes（做 UIA 自动化必需）、pywin32（可选，增强稳定性）

用法
----
    python testo184_vi2_export.py            # 导出当前所有已插入设备
    python testo184_vi2_export.py --watch    # 持续监视，插入新设备自动导出
    python testo184_vi2_export.py --force    # 已存在也重新导出
    python testo184_vi2_export.py --base "D:\X"   # 指定输出根目录
    python testo184_vi2_export.py --selftest # 只做环境自检，不动 Comsoft

输出：<输出根>\<YYYY-MM-DD>\<设备号>.vi2 + 采集日志.csv
退出码：0 正常；2 有设备导出失败；1 环境错误
"""

import argparse
import csv
import ctypes
import ctypes.wintypes as wt
import json
import os
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

TOOL_DIR = Path(__file__).resolve().parent
# PyInstaller 打包后：config/日志写到 exe 同目录（可写），而非 _MEIPASS 临时目录
if getattr(sys, 'frozen', False):
    import sys as _sys
    _exe_dir = Path(_sys.executable).resolve().parent
    if _exe_dir.is_dir() and os.access(str(_exe_dir), os.W_OK):
        TOOL_DIR = _exe_dir

DEFAULT_CONFIG = {
    "output_dir": "",
    "poll_interval": 3.0,
    "date_format": "%Y-%m-%d",
    "comsoft_exe": r"D:\Testo\Comfort Software\cc4.exe",
    "close_tabs_after_export": True,
}
TREE_ITEM_RE = re.compile(r"testo\s*184.*?(\d{6,})", re.IGNORECASE)
REPORT_LEAF = "testo 184 measurement report.pdf"
MAIN_CLASS_PREFIX = "Afx:400000"

u32 = ctypes.windll.user32
k32 = ctypes.windll.kernel32

# WM_* / BM_* 常量
WM_SETTEXT = 0x000C
WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E
WM_SYSCOMMAND = 0x0112
WM_CLOSE = 0x0010
SC_CLOSE = 0xF060
BM_CLICK = 0x00F5
SW_RESTORE = 9

try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass


def log(msg: str = ""):
    print(msg, flush=True)


# ============================================================
#  依赖 / 自检
# ============================================================
def load_config() -> dict:
    cfg = dict(DEFAULT_CONFIG)
    p = TOOL_DIR / "config.json"
    if p.exists():
        try:
            import json
            txt = p.read_text(encoding="utf-8-sig")
            try:
                data = json.loads(txt)
            except json.JSONDecodeError:
                # 容错：Windows 路径里常见的单反斜杠（"C:\Users"）→ 转义
                fixed = re.sub(r'(?<!\\)\\(?!["\\/bfnrtu])', r'\\\\', txt)
                data = json.loads(fixed)
            if isinstance(data, dict):
                cfg.update({k: v for k, v in data.items() if k in DEFAULT_CONFIG})
        except Exception as e:
            log(f"[警告] config.json 读取失败，已用默认配置：{e}")
    return cfg


def check_deps():
    """返回 (UIA 模块, 是否可用)。不导入 pywinauto 的 uia 后端。"""
    try:
        import comtypes.client  # noqa: F401
    except ImportError:
        return None
    try:
        import comtypes.gen.UIAutomationClient as UIA
        return UIA
    except ImportError:
        try:
            import comtypes.client as cc
            cc.GetModule("UIAutomationCore.dll")
            import comtypes.gen.UIAutomationClient as UIA
            return UIA
        except Exception:
            return None


# ============================================================
#  UIA 轻封装（RawViewWalker 逐层遍历，绝不 FindAll）
# ============================================================
class UIA:
    def __init__(self):
        import comtypes.client as cc
        import comtypes.gen.UIAutomationClient as nv
        self.n = nv
        self.api = cc.CreateObject("{ff48dba4-60ef-4201-aa87-54103eef594e}",
                                   interface=nv.IUIAutomation)
        self.root = self.api.GetRootElement()
        self.walker = self.api.RawViewWalker
        self.TS_DESC = nv.TreeScope_Descendants

    # ---- 基础属性（全部做异常保护，叶子节点会抛 NULL COM pointer ----）
    @staticmethod
    def name(el) -> str:
        try:
            return el.CurrentName or ""
        except Exception:
            return ""

    @staticmethod
    def cls(el) -> str:
        try:
            return el.CurrentClassName or ""
        except Exception:
            return ""

    @staticmethod
    def ctype(el) -> int:
        try:
            return int(el.CurrentControlType)
        except Exception:
            return 0

    @staticmethod
    def hwnd(el) -> int:
        try:
            return int(el.CurrentNativeWindowHandle or 0)
        except Exception:
            return 0

    @staticmethod
    def rect(el):
        try:
            r = el.CurrentBoundingRectangle
            return (int(r.left), int(r.top), int(r.right), int(r.bottom))
        except Exception:
            return None

    def children(self, el):
        """按 RawViewWalker 取直接子元素（替代会死锁的 FindAll）。"""
        out = []
        try:
            c = self.walker.GetFirstChildElement(el)
        except Exception:
            return out
        n = 0
        while c is not None and n < 500:
            out.append(c)
            try:
                c = self.walker.GetNextSiblingElement(c)
            except Exception:
                break
            n += 1
        return out

    def walk(self, el, max_depth=14):
        """深度优先遍历，yield (element, depth)。"""
        stack = [(el, 0)]
        seen = 0
        while stack:
            e, d = stack.pop()
            yield e, d
            seen += 1
            if seen > 6000 or d >= max_depth:
                continue
            kids = self.children(e)
            for k in reversed(kids):
                stack.append((k, d + 1))

    def collect(self, el, pred, max_depth=14):
        return [e for e, _ in self.walk(el, max_depth) if pred(e)]


# ============================================================
#  win32 底层动作
# ============================================================
def foreground(hwnd: int):
    """把窗口激活到前台。

    先直接判定：若已是前台立即返回（绝大多数情况 0ms）。
    否则 SetForegroundWindow 后高频轮询，最多 0.35s。
    """
    try:
        if u32.GetForegroundWindow() == hwnd:
            return
        u32.ShowWindow(hwnd, SW_RESTORE)
        u32.SetForegroundWindow(hwnd)
        deadline = time.time() + 0.35
        while time.time() < deadline:
            if u32.GetForegroundWindow() == hwnd:
                return
            time.sleep(0.01)
    except Exception:
        pass
        pass


def find_main_hwnd(pid: int) -> int:
    """用 win32 EnumWindows 找 Comsoft 主窗口（~11ms）。

    **不用 UIA root.FindFirst**：该调用在本机会因 UIA 提供者无响应而
    永久死锁（Ctrl+C 也无效）。win32 枚举瞬间返回，且非常稳。
    匹配条件：属该 pid + 可见 + 类名以 Afx:400000 开头 + 标题以 Testo 开头。
    """
    hits = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wt.HWND, wt.LPARAM)
    def cb(h, lp):
        p = wt.DWORD()
        u32.GetWindowThreadProcessId(h, ctypes.byref(p))
        if p.value != pid or not u32.IsWindowVisible(h):
            return True
        cls = ctypes.create_unicode_buffer(256)
        u32.GetClassNameW(h, cls, 256)
        if not cls.value.startswith(MAIN_CLASS_PREFIX):
            return True
        t = ctypes.create_unicode_buffer(512)
        u32.GetWindowTextW(h, t, 512)
        if "Testo" not in t.value:
            return True
        hits.append(h)
        return True

    u32.EnumWindows(cb, 0)
    return hits[0] if hits else 0


def enum_doc_windows(hwnd: int):
    """枚举主窗口下 AfxFrameOrView42 文档视图窗口 (hwnd, title)。

    排除 "表" / "图表" 两个固定子视图，只留真正的测量报告文档。
    """
    out = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wt.HWND, wt.LPARAM)
    def cb(ch, lp):
        cls = ctypes.create_unicode_buffer(64)
        u32.GetClassNameW(ch, cls, 64)
        if cls.value == "AfxFrameOrView42":
            t = ctypes.create_unicode_buffer(256)
            u32.GetWindowTextW(ch, t, 256)
            if t.value and t.value not in ("表", "图表"):
                out.append((ch, t.value))
        return True

    u32.EnumChildWindows(hwnd, cb, 0)
    return out


def enum_doc_titles(hwnd: int):
    return [t for _, t in enum_doc_windows(hwnd)]


def _doc_frame(doc_view: int) -> int:
    """文档视图(AfxFrameOrView42)所属的 MDI 子框架窗口。

    MDI 文档的标题栏（含右上角 × ）属于子框架窗口，而不是视图本身。
    向上找：视图 → ... → 子框架（父窗口是 MDIClient）。
    """
    p = u32.GetParent(doc_view)
    for _ in range(6):
        if not p:
            return 0
        parent = u32.GetParent(p)
        if not parent:
            return p
        pc = ctypes.create_unicode_buffer(64)
        u32.GetClassNameW(parent, pc, 64)
        if pc.value == "MDIClient":
            return p          # p 即子框架窗口（其父是 MDIClient）
        p = parent
    return p


def _click_doc_close(doc_view: int) -> bool:
    """点击文档子窗口标题栏右上角的 × 按钮（物理点击）。

    × 位于子框架窗口标题栏最右端：窗口右边 - 约 一个按钮宽度 处。
    用 GetSystemMetrics(SM_CXSIZE) 估算按钮宽度，纵向取标题栏中线。
    这是"点文档右上角关闭"的安全做法，绝不会误关主程序。
    """
    frame = _doc_frame(doc_view)
    if not frame or not u32.IsWindow(frame):
        return False
    r = wt.RECT()
    u32.GetWindowRect(frame, ctypes.byref(r))
    if r.right <= r.left or r.bottom <= r.top:
        return False
    SM_CXSIZE = u32.GetSystemMetrics(30)    # 标题栏按钮宽度
    SM_CYSIZE = u32.GetSystemMetrics(31)    # 标题栏按钮高度
    SM_CYSMCAPTION = u32.GetSystemMetrics(51)  # 小标题栏高度
    bw = max(SM_CXSIZE, 16)
    bh = max(SM_CYSIZE, 16)
    # 小标题栏：× 中心大致在 (right - bw/2, top + bh/2 + 1)
    cx = r.right - bw // 2 - 1
    cy = r.top + max(bh, SM_CYSMCAPTION) // 2
    phys_click(cx, cy)
    return True


def close_docs(hwnd: int, max_rounds=20, pid: int = 0) -> int:
    """逐个关闭文档 —— **等价于点文档右上角的 ×**，绝不关主程序。

    安全要点（关键！）：
    * **绝不**给主窗口发 SC_CLOSE —— 没有文档时那会直接关掉整个 Comsoft！
    * 首选：对文档视图发 WM_CLOSE（关闭该文档，一次一个）。
    * 兜底：物理点击该文档子窗口标题栏的 × 按钮。
    * 每次关闭后**立刻重取**文档列表（关闭会改变窗口集合）。
    * **等待中内联应答确认框**：WM_CLOSE 会异步触发「将改动保存到
      ComsoftN？」模态框，不点掉它文档永远关不掉（曾致 21s 空转 +
      后续所有设备打不开）。一律点"否"——.vi2 在导出流程里已另存。
    """
    for _ in range(max_rounds):
        docs = enum_doc_windows(hwnd)
        if not docs:
            break
        n = len(docs)
        doc_h, _title = docs[0]
        # 策略1：给文档视图发 WM_CLOSE（等价文档关闭，最快）
        u32.PostMessageW(doc_h, WM_CLOSE, 0, 0)
        # 高频轮询：文档数下降即进入下一轮；弹出确认框立即点"否"
        # 注意：点"否"后文档异步关闭需要几百毫秒 —— 应答过就把等待窗口
        # 延长，避免超时后误去点**下一个**文档的 ×（会引发连环确认框）。
        deadline = time.time() + 0.5
        while time.time() < deadline:
            if len(enum_doc_windows(hwnd)) < n:
                break
            if pid and answer_all_confirms(pid, yes=False, max_rounds=2):
                deadline = max(deadline, time.time() + 0.6)
                continue
            time.sleep(0.02)
        # 策略2：WM_CLOSE 没生效 → 物理点该文档右上角 ×
        if len(enum_doc_windows(hwnd)) >= n:
            u32.SetForegroundWindow(hwnd)
            _click_doc_close(doc_h)
            deadline = time.time() + 0.5
            while time.time() < deadline:
                if len(enum_doc_windows(hwnd)) < n:
                    break
                if pid and answer_all_confirms(pid, yes=False, max_rounds=2):
                    deadline = max(deadline, time.time() + 0.6)
                    continue
                time.sleep(0.02)
    return len(enum_doc_windows(hwnd))


def phys_click(x: int, y: int, double=False):
    """物理点击（可选双击）。

    时序说明：SetCursorPos 后给 15ms 让光标就位；双击间隔 30ms，
    远小于系统双击阈值（默认 500ms），不会被识别成两次单击。
    """
    u32.SetCursorPos(int(x), int(y))
    time.sleep(0.015)
    times = 2 if double else 1
    for i in range(times):
        u32.mouse_event(0x0002, 0, 0, 0, 0)  # LEFTDOWN
        time.sleep(0.02)
        u32.mouse_event(0x0004, 0, 0, 0, 0)  # LEFTUP
        if i + 1 < times:
            time.sleep(0.03)


def click_element(uia: UIA, el, double=False) -> bool:
    r = uia.rect(el)
    if not r:
        return False
    cx, cy = (r[0] + r[2]) // 2, (r[1] + r[3]) // 2
    phys_click(cx, cy, double)
    return True


def dialog_hwnd(title: str) -> int:
    return u32.FindWindowW("#32770", title)


def any_dialog_titles(pid: int):
    """枚举该 pid 下所有可见 #32770 对话框标题。"""
    found = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wt.HWND)
    def cb(h):
        if not u32.IsWindowVisible(h):
            return True
        p = wt.DWORD()
        u32.GetWindowThreadProcessId(h, ctypes.byref(p))
        if p.value != pid:
            return True
        cls = ctypes.create_unicode_buffer(256)
        u32.GetClassNameW(h, cls, 256)
        if cls.value == "#32770":
            t = ctypes.create_unicode_buffer(512)
            u32.GetWindowTextW(h, t, 512)
            if t.value:
                found.append((h, t.value))
        return True

    u32.EnumWindows(cb, 0)
    return found


def _with_attached(hwnd, fn):
    """在附加输入队列的情况下执行 fn（跨进程文本/焦点操作必需）。"""
    my = k32.GetCurrentThreadId()
    tgt = u32.GetWindowThreadProcessId(hwnd, None)
    ok = False
    try:
        if tgt:
            ok = bool(u32.AttachThreadInput(my, tgt, True))
        return fn()
    finally:
        if ok:
            try:
                u32.AttachThreadInput(my, tgt, False)
            except Exception:
                pass


def set_edit_text(hwnd_dlg: int, ctrl_id: int, text: str) -> bool:
    """给对话框 Edit 写入文本，并**回读确认**写入成功。

    该 MFC 对话框的 Edit 只接受 AttachThreadInput 后的 WM_SETTEXT，
    且必须在对话框完全初始化后；故采用"写入 → 回读校验 → 不符则重试"。
    """
    edit = u32.GetDlgItem(hwnd_dlg, ctrl_id)
    if not edit:
        return False

    def do():
        u32.SetForegroundWindow(hwnd_dlg)
        u32.SetFocus(edit)
        u32.SendMessageW(edit, WM_SETTEXT, 0, ctypes.c_wchar_p(text))
        # 高频轮询回读（不做固定 sleep），通常 1~2 次即命中
        for _ in range(12):
            ln = u32.SendMessageW(edit, WM_GETTEXTLENGTH, 0, 0)
            buf = ctypes.create_unicode_buffer(1024)
            u32.SendMessageW(edit, WM_GETTEXT, 1024, ctypes.byref(buf))
            if ln > 0 and buf.value.strip() == text.strip():
                return True
            time.sleep(0.01)
        return False

    for attempt in range(4):
        if _with_attached(hwnd_dlg, do):
            return True
        time.sleep(0.15)
    return False


def click_dialog_button(hwnd_dlg: int, ctrl_id: int) -> bool:
    btn = u32.GetDlgItem(hwnd_dlg, ctrl_id)
    if not btn:
        return False

    def do():
        u32.SendMessageW(btn, BM_CLICK, 0, 0)
        return True

    _with_attached(hwnd_dlg, do)
    # 等对话框关闭（BM_CLICK 生效即走，最多 0.2s），未关再补物理点击
    deadline = time.time() + 0.2
    while time.time() < deadline:
        if not u32.IsWindowVisible(hwnd_dlg):
            return True
        time.sleep(0.02)
    if u32.IsWindow(hwnd_dlg) and u32.IsWindowVisible(hwnd_dlg):
        r = wt.RECT()
        u32.GetWindowRect(btn, ctypes.byref(r))
        if r.right > r.left and r.bottom > r.top:
            phys_click((r.left + r.right) // 2, (r.top + r.bottom) // 2)
    return True


def _enum_child_windows(hwnd):
    """枚举子窗口 (hwnd, class, ctrl_id, text)。"""
    out = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wt.HWND, wt.LPARAM)
    def cb(ch, lp):
        t = ctypes.create_unicode_buffer(128)
        u32.GetWindowTextW(ch, t, 128)
        c = ctypes.create_unicode_buffer(128)
        u32.GetClassNameW(ch, c, 128)
        out.append((ch, c.value, u32.GetDlgCtrlID(ch), t.value))
        return True

    u32.EnumChildWindows(hwnd, cb, 0)
    return out


def answer_confirm_dialog(hwnd_dlg: int, yes: bool) -> bool:
    """覆盖确认框（TaskDialog 自绘按钮，control_id 均为 0）。

    这类对话框的按钮不是标准 GetDlgItem 子控件，必须用 EnumChildWindows
    遍历、按按钮文本匹配，再物理点击其中心坐标。
    重要：点击前必须先把该对话框激活到前台，否则点击不生效。
    """
    # 先激活对话框（SetForegroundWindow 即刻生效，无需固定 sleep）
    try:
        u32.SetForegroundWindow(hwnd_dlg)
    except Exception:
        pass

    kids = _enum_child_windows(hwnd_dlg)
    want = "是(&Y)" if yes else "否(&N)"

    def _hit(ch):
        r = wt.RECT()
        u32.GetWindowRect(ch, ctypes.byref(r))
        phys_click((r.left + r.right) // 2, (r.top + r.bottom) // 2)

    # 1) 精确文本匹配
    for ch, cls, cid, txt in kids:
        if cls == "Button" and txt.strip() == want:
            _hit(ch)
            return True

    # 2) 宽松匹配（是否/Yes/No）
    keys = ("是", "Yes") if yes else ("否", "No")
    for ch, cls, cid, txt in kids:
        if cls == "Button" and any(k in txt for k in keys):
            _hit(ch)
            return True

    # 3) 标准 MsgBox 兜底（id 6=是 7=否）
    std = 6 if yes else 7
    h = u32.GetDlgItem(hwnd_dlg, std)
    if h:
        _with_attached(hwnd_dlg,
                       lambda: (u32.SendMessageW(h, BM_CLICK, 0, 0), True)[1])
        return True
    return False


def answer_all_confirms(pid: int, yes: bool = False, max_rounds: int = 6) -> int:
    """循环应答该进程所有 是/否 确认框，直到一个不剩。

    关键场景：关闭文档时 Comsoft 异步弹出「将改动保存到 ComsoftN？」
    （标题为 cc4 的 #32770，按钮 是(&Y)/否(&N)/取消）。该框是模态的，
    **不点掉它整个界面全部失灵**——双击无效、文档关不掉。
    必须在等待循环里内联应答，而不能等循环结束后再兜底。
    """
    n = 0
    for _ in range(max_rounds):
        hit = False
        for h, title in any_dialog_titles(pid):
            if title in ("保存为", "另存为"):
                continue
            if answer_confirm_dialog(h, yes=yes):
                n += 1
                hit = True
        if not hit:
            break
        time.sleep(0.06)   # 等下一个框弹出（Comsoft 可能连弹多个）
    return n


def dismiss_leftover_dialogs(pid: int):
    """兜底：取消残留的 保存为 / 关闭所有确认框，避免卡死 Comsoft。"""
    n = 0
    for h, title in any_dialog_titles(pid):
        try:
            if title in ("保存为", "另存为"):
                click_dialog_button(h, 2)  # 取消
                n += 1
            else:
                # 任意确认框：优先选“否”，避免覆盖/丢弃
                if not answer_confirm_dialog(h, yes=False):
                    # 没有标准按钮的，尝试 ESC
                    _with_attached(h, lambda: u32.PostMessageW(h, 0x0100, 0x1B, 0))
                n += 1
        except Exception:
            continue
    return n


# ============================================================
#  Comsoft 控制
# ============================================================
UI_CACHE_FILE = Path(__file__).with_name("ui_cache.json")


def _load_ui_cache() -> dict:
    """磁盘级 UI 坐标缓存（窗口相对偏移，跨运行复用，首次省 ~1.2s UIA 遍历）。"""
    try:
        d = json.loads(UI_CACHE_FILE.read_text(encoding="utf-8"))
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def _save_ui_cache(**kv):
    try:
        d = _load_ui_cache()
        d.update(kv)
        UI_CACHE_FILE.write_text(
            json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
    except Exception:
        pass


class Comsoft:
    def __init__(self):
        self.uia = UIA()
        self.pid = 0
        self.main = None
        self._hwnd = 0
        # "另存为(A)..." 菜单项坐标缓存（同一窗口会话内稳定）
        self._saveas_xy = None

    def running_pid(self) -> int:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq cc4.exe"],
                             capture_output=True, text=True,
                             encoding="gbk", errors="replace").stdout
        m = re.search(r"cc4\.exe\s+(\d+)", out)
        return int(m.group(1)) if m else 0

    def connect(self, timeout=25.0) -> bool:
        """找 Comsoft 主窗口。

        **关键：用 win32 EnumWindows（~11ms），不用 UIA root.FindFirst**
        （后者在本机会永久死锁）。拿到 hwnd 后由它换 UIA 元素（ElementFromHandle，
        ~76ms，安全），供后续 devices()/菜单定位使用。
        """
        self.pid = self.running_pid()
        if not self.pid:
            return False
        deadline = time.time() + timeout
        while time.time() < deadline:
            h = find_main_hwnd(self.pid)
            if h:
                if h != self._hwnd:
                    self._saveas_xy = None   # 新窗口会话 → 清菜单坐标缓存
                self._hwnd = h
                try:
                    self.main = self.uia.api.ElementFromHandle(h)
                except Exception:
                    self.main = None
                return True
            time.sleep(0.25)
        return False

    @property
    def hwnd(self) -> int:
        """主窗口 hwnd（缓存，win32 直取，1ms 内）。"""
        h = getattr(self, "_hwnd", 0)
        if h and u32.IsWindow(h):
            return h
        if self.main is not None:
            h = self.uia.hwnd(self.main)
            if h:
                self._hwnd = h
                return h
        h = find_main_hwnd(self.pid)
        self._hwnd = h
        return h

    def foreground(self):
        foreground(self.hwnd)

    # ---- 部件定位 ----
    def devices(self, use_cache=False):
        """{serial: element}；存档树只显示当前插着的设备。

        全树遍历约 0.5s，故支持缓存（监视模式下按轮询周期复用）。
        """
        if use_cache:
            c = getattr(self, "_dev_cache", None)
            if c and c.get("hwnd") == self.hwnd and time.time() < c.get("exp", 0):
                return c["map"]
        out = {}
        # 设备节点实测在深度 5~6 —— 限深遍历，比全树快约一半
        for el, _d in self.uia.walk(self.main, max_depth=8):
            nm = self.uia.name(el)
            if not nm:
                continue
            m = TREE_ITEM_RE.search(nm)
            if not m or REPORT_LEAF in nm:
                continue
            if self.uia.cls(el) == "SysTreeView32":
                continue
            serial = m.group(1)
            out.setdefault(serial, el)
        self._dev_cache = {"hwnd": self.hwnd, "map": out, "exp": time.time() + 3.0}
        return out

    def report_leaf(self, dev_el):
        """设备节点下的 measurement report 子节点。"""
        hits = self.uia.collect(dev_el, lambda e: REPORT_LEAF in self.uia.name(e), 3)
        if hits:
            return hits[0]
        # 未展开 → 尝试展开后轮询
        try:
            dev_el.GetCurrentPattern(self.uia.n.UIA_ExpandCollapsePatternId).Expand()
        except Exception:
            pass
        deadline = time.time() + 1.5
        while time.time() < deadline:
            hits = self.uia.collect(dev_el, lambda e: REPORT_LEAF in self.uia.name(e), 3)
            if hits:
                return hits[0]
            time.sleep(0.05)
        return None

    def menu_button(self, label: str):
        """功能区按钮定位（UIA 元素）；结果缓存，避免每次全树遍历。

        缓存键含主窗口句柄，Comsoft 重启后自动失效。
        """
        cache = getattr(self, "_menu_cache", None)
        if cache is None or cache.get("hwnd") != self.hwnd:
            cache = {"hwnd": self.hwnd, "map": {}}
            self._menu_cache = cache
        if label in cache["map"]:
            return cache["map"][label]
        hit = None
        # 实测"文件"在深度 3 —— 限深遍历比全树（1.4s）快一半以上
        for el, _d in self.uia.walk(self.main, max_depth=5):
            if self.uia.name(el) == label and self.uia.ctype(el) in (50000, 50011, 50019):
                hit = el
                break
        cache["map"][label] = hit
        return hit

    def menu_button_xy(self, label: str):
        """功能区按钮中心坐标；带坐标缓存（点击用，最快）。

        缓存的是坐标而非元素对象 —— 避免元素失效，且省去 rect 查询。
        """
        cache = getattr(self, "_menu_xy_cache", None)
        if cache is None or cache.get("hwnd") != self.hwnd:
            cache = {"hwnd": self.hwnd, "map": {}}
            self._menu_xy_cache = cache
        if label in cache["map"]:
            return cache["map"][label]
        el = self.menu_button(label)
        xy = None
        if el is not None:
            r = self.uia.rect(el)
            if r:
                xy = ((r[0] + r[2]) // 2, (r[1] + r[3]) // 2)
        cache["map"][label] = xy
        return xy

    def click_menu(self, label: str) -> bool:
        """按坐标点击功能区按钮（比 click_element 少一次 rect 查询）。"""
        xy = self.menu_button_xy(label)
        if not xy:
            return False
        phys_click(xy[0], xy[1])
        return True

    def doc_titles(self):
        """当前打开的文档标签标题（顶层视图；排除 表/图表 子视图）。

        走 win32 EnumChildWindows（~0.1ms），**不用 UIA 全树遍历（~537ms）**。
        这是性能关键：本方法在高频轮询中被反复调用。
        """
        return enum_doc_titles(self.hwnd)

    def doc_count(self) -> int:
        return len(self.doc_titles())

    # ---- 动作 ----
    def open_report(self, dev_el, retries=2) -> bool:
        """双击 report 节点打开文档。

        **只双击一遍**（实测单次双击 108ms 即打开；此前双击两遍 + 重试纯属浪费）。
        双击前必须激活窗口，否则无效。
        若 0.8s 内未出现新文档，才做一次兜底重试（防止首次点击丢失）。
        """
        for attempt in range(1, retries + 1):
            # 双击前先清场：任何残留模态框（如「将改动保存到？」）都会让双击失效
            answer_all_confirms(self.pid, yes=False, max_rounds=2)
            leaf = self.report_leaf(dev_el)
            if leaf is None:
                if attempt < retries:
                    time.sleep(0.15)
                    continue
                return False
            n0 = self.doc_count()
            self.foreground()          # 关键：不激活窗口时双击无效
            if not click_element(self.uia, leaf, double=True):
                time.sleep(0.15)
                continue
            deadline = time.time() + 0.8
            while time.time() < deadline:
                if self.doc_count() > n0:
                    return True
                time.sleep(0.02)
            # 未打开 → 极少数情况下焦点丢失，快速再试一次
            self.foreground()
            time.sleep(0.08)
        return False

    def _wait_popup_bar(self, timeout=0.4) -> bool:
        """等 Codejock 菜单弹层窗口（XTPPopupBar）出现（win32 高频轮询）。

        实测点[文件]后 ~46ms 弹层即出现；出现后菜单项坐标才可点。
        """
        deadline = time.time() + timeout
        while time.time() < deadline:
            h = u32.FindWindowExW(None, 0, "XTPPopupBar", None)
            if h:
                p = wt.DWORD()
                u32.GetWindowThreadProcessId(h, ctypes.byref(p))
                if p.value == self.pid:
                    return True
            time.sleep(0.01)
        return False

    def _win_geo(self):
        """主窗口 (left, top, width, height)。"""
        r = wt.RECT()
        u32.GetWindowRect(self.hwnd, ctypes.byref(r))
        return r.left, r.top, r.right - r.left, r.bottom - r.top

    def _file_btn_xy(self):
        """[文件]按钮坐标：会话缓存 → 磁盘缓存（窗口尺寸一致才信任）→ UIA 定位。"""
        cache = getattr(self, "_menu_xy_cache", None)
        if cache and cache.get("hwnd") == self.hwnd and "文件" in cache["map"]:
            return cache["map"]["文件"]
        L, T, W, H = self._win_geo()
        uc = _load_ui_cache()
        if (uc.get("win") == [W, H] and isinstance(uc.get("file_btn"), list)
                and len(uc["file_btn"]) == 2):
            try:
                xy = (L + int(uc["file_btn"][0]), T + int(uc["file_btn"][1]))
                self._menu_xy_cache = {"hwnd": self.hwnd, "map": {"文件": xy}}
                return xy
            except (TypeError, ValueError):
                pass
        xy = self.menu_button_xy("文件")     # UIA 限深遍历（最慢路径）
        if xy:
            _save_ui_cache(win=[W, H], file_btn=[xy[0] - L, xy[1] - T])
        return xy

    def save_as(self, timeout=10.0):
        """文件 → 另存为(A)...；返回出现的保存对话框 hwnd（或 0）。

        性能要点（实测 2026-10-08/09）：
        * "文件"按钮：磁盘缓存窗口相对偏移 → 会话缓存 → UIA 限深定位，
          首台也免 UIA 遍历（省 ~1.2s）。
        * 点[文件]后轮询 XTPPopupBar 弹层（~50ms）代替固定 sleep。
        * "另存为(A)..." 菜单项坐标非常稳定，直接按缓存坐标点击；
          同样走 磁盘 → 会话 → UIA 三级缓存。
        * 任一缓存点击未奏效 → ESC 关菜单 → 回退 UIA 定位（自愈）。
        """
        fbtn = self._file_btn_xy()
        if not fbtn:
            return 0
        answer_all_confirms(self.pid, yes=False, max_rounds=2)  # 模态框会让菜单点不开
        self.foreground()          # 点菜单前需激活窗口
        phys_click(fbtn[0], fbtn[1])

        # 另存为菜单坐标：会话缓存 → 磁盘缓存
        xy = self._saveas_xy
        if xy is None:
            L, T, W, H = self._win_geo()
            uc = _load_ui_cache()
            if (uc.get("win") == [W, H] and isinstance(uc.get("saveas_menu"), list)
                    and len(uc["saveas_menu"]) == 2):
                try:
                    self._saveas_xy = (L + int(uc["saveas_menu"][0]),
                                       T + int(uc["saveas_menu"][1]))
                except (TypeError, ValueError):
                    pass
            xy = self._saveas_xy

        # 先试缓存坐标（最快路径）
        if xy:
            # 等 Codejock 菜单弹层（XTPPopupBar）出现，实测 ~50ms
            if not self._wait_popup_bar(0.4):
                time.sleep(0.2)
            phys_click(xy[0], xy[1])
            deadline = time.time() + 1.2
            while time.time() < deadline:
                h = dialog_hwnd("保存为")
                if h:
                    return h
                time.sleep(0.02)
            # 缓存失效 → 关掉可能残留的菜单，走回退路径
            u32.keybd_event(0x1B, 0, 0, 0)
            u32.keybd_event(0x1B, 0, 2, 0)
            time.sleep(0.15)
            phys_click(fbtn[0], fbtn[1])

        # 回退：UIA 浅遍历定位"另存为"（深度 2 即可命中）
        items = []
        deadline = time.time() + 4.0
        while time.time() < deadline:
            items = [e for e, d in self.uia.walk(self.main, max_depth=3)
                     if self.uia.name(e).startswith("另存为")]
            if items:
                break
            time.sleep(0.02)
        if not items:
            u32.keybd_event(0x1B, 0, 0, 0)  # ESC
            u32.keybd_event(0x1B, 0, 2, 0)
            return 0
        r = self.uia.rect(items[0])
        if r:
            self._saveas_xy = ((r[0] + r[2]) // 2, (r[1] + r[3]) // 2)  # 记入缓存
            L, T, W, H = self._win_geo()
            _save_ui_cache(win=[W, H],
                           saveas_menu=[self._saveas_xy[0] - L, self._saveas_xy[1] - T])
        if not click_element(self.uia, items[0]):
            return 0

        deadline = time.time() + timeout
        while time.time() < deadline:
            h = dialog_hwnd("保存为")
            if h:
                return h
            time.sleep(0.03)
        return 0

    def close_tabs(self, max_rounds=20) -> int:
        """关闭所有文档（导出后释放文件句柄）。

        点文档右上角 ×（等价手动关闭），绝不发 SC_CLOSE 给主窗口。
        结束前**循环应答**残留确认框（「将改动保存到？」可能晚于
        WM_CLOSE 数百毫秒才弹出，且可能连弹多个），确保一个不剩。
        """
        n = close_docs(self.hwnd, max_rounds, pid=self.pid)
        # 兜底：循环应答残留确认框，直到干净（避免模态框卡死后续流程）
        answer_all_confirms(self.pid, yes=False, max_rounds=8)
        return len(enum_doc_windows(self.hwnd))


# ============================================================
#  导出单个设备
# ============================================================
def serial_in_file(path: Path, serial: str) -> bool:
    try:
        return b"SerialNumber\t" + serial.encode() in path.read_bytes()
    except OSError:
        return False


def wait_file_ready(path: Path, timeout=20.0, settle=0.12) -> bool:
    """等文件写完（大小连续稳定 settle 秒且非空，或内部已含 SerialNumber）。

    Comsoft 保存时文件会先以 0 字节出现、随后写入，故必须等内容落盘再校验。
    """
    deadline = time.time() + timeout
    last = -1
    stable_since = None
    while time.time() < deadline:
        try:
            size = path.stat().st_size if path.exists() else 0
        except OSError:
            size = 0
        if size > 0:
            # 若已含关键标记，直接成功（最快路径）
            try:
                if b"SerialNumber\t" in path.read_bytes():
                    return True
            except OSError:
                pass
            if size == last:
                if stable_since is None:
                    stable_since = time.time()
                elif time.time() - stable_since >= settle:
                    return True
            else:
                last = size
                stable_since = None
        time.sleep(0.05)
    return path.exists() and path.stat().st_size > 0


def export_one(cs: Comsoft, serial: str, dev_el, target: Path, force: bool) -> str:
    if target.exists() and not force:
        if serial_in_file(target, serial):
            log(f"[{serial}] 已存在且校验一致，跳过（--force 可重导）")
            return "跳过"

    # 保证干净起点：关掉所有已打开文档，避免"另存为"继承上一个文档的路径而弹覆盖框
    if cs.doc_count():
        cs.close_tabs()

    log(f"[{serial}] 打开测量报告...")
    if not cs.open_report(dev_el):
        log(f"[{serial}] 未能打开文档（存档树节点异常？）")
        dismiss_leftover_dialogs(cs.pid)
        return "打开失败"

    # 先删掉旧目标文件（在弹"另存为"之前），这样 Comsoft 检测不到同名文件，
    # 就不会弹覆盖确认框 —— 少一次交互，也少一次时序风险
    if target.exists():
        try:
            target.unlink()
        except OSError:
            pass

    log(f"[{serial}] 文件 → 另存为...")
    hdlg = cs.save_as()
    if not hdlg:
        log(f"[{serial}] 未出现[保存为]对话框（需中文界面 Comsoft）")
        dismiss_leftover_dialogs(cs.pid)
        return "无保存对话框"

    # 等对话框完全就绪（Edit 可写），否则填名会静默失败。
    # 实测：必须确认 Edit 存在**且可见、已启用**，否则快速路径下可能
    # 抓到尚未初始化完的对话框，导致写入被丢弃、Comsoft 用默认名保存。
    deadline = time.time() + 3.0
    while time.time() < deadline:
        e = u32.GetDlgItem(hdlg, 1152)
        if e and u32.IsWindow(e) and u32.IsWindowVisible(e) and u32.IsWindowEnabled(e):
            break
        time.sleep(0.02)
    time.sleep(0.05)   # 极小稳定期，确保对话框消息循环就绪

    log(f"[{serial}] 填写文件名并保存：{target.name}")
    if not set_edit_text(hdlg, 1152, str(target)):
        log(f"[{serial}] 无法写入文件名（对话框异常）")
        click_dialog_button(hdlg, 2)
        return "填写文件名失败"

    # 二次确认：写入后再回读一次，避免"读到旧值但实际已被清空"
    e2 = u32.GetDlgItem(hdlg, 1152)
    buf2 = ctypes.create_unicode_buffer(1024)
    u32.SendMessageW(e2, WM_GETTEXT, 1024, ctypes.byref(buf2))
    if buf2.value.strip() != str(target).strip():
        log(f"[{serial}] 文件名回读不符，重写一次：{buf2.value!r}")
        if not set_edit_text(hdlg, 1152, str(target)):
            click_dialog_button(hdlg, 2)
            return "填写文件名失败"

    click_dialog_button(hdlg, 1)   # 保存

    # 等待保存完成：单层紧凑循环，同时处理覆盖确认框与文件落盘
    # 注意：点"保存"后对话框关闭到文件创建之间有短暂空窗期，
    #       不能用"对话框没了且文件还没出现"立即判失败，必须给宽限期。
    start = time.time()
    deadline = start + 30.0
    overwrite_answered = False
    ready = False
    last_size = -1
    stable_since = None
    while time.time() < deadline:
        # 1) 覆盖确认框优先处理
        for h, t in any_dialog_titles(cs.pid):
            if t == "保存为":
                continue
            if not overwrite_answered:
                answer_confirm_dialog(h, yes=True)
                overwrite_answered = True
                log(f"[{serial}] 出现覆盖确认框，已选[是]")

        # 2) 文件是否就绪（含关键标记即成功，最快路径）
        try:
            size = target.stat().st_size if target.exists() else 0
        except OSError:
            size = 0
        if size > 0:
            try:
                if b"SerialNumber\t" in target.read_bytes():
                    ready = True
                    break
            except OSError:
                pass
            if size == last_size:
                if stable_since is None:
                    stable_since = time.time()
                elif time.time() - stable_since >= 0.15:
                    ready = True
                    break
            else:
                last_size = size
                stable_since = None

        # 3) 失败判定：只在给了 2 秒宽限、且确实无任何对话框、文件也不存在时
        if (time.time() - start > 2.0
                and not dialog_hwnd("保存为")
                and not any_dialog_titles(cs.pid)
                and not target.exists()):
            break
        time.sleep(0.04)

    # 兜底：若确认框仍在（未被识别），强制清理后再等一次
    if not ready and any_dialog_titles(cs.pid):
        dismiss_leftover_dialogs(cs.pid)
        time.sleep(0.2)
        if wait_file_ready(target, timeout=3.0):
            ready = True

    if ready and serial_in_file(target, serial):
        log(f"[{serial}] OK 已保存 {target.name}"
            f"（{target.stat().st_size} 字节，设备号校验一致）")
        return "OK"
    if ready:
        log(f"[{serial}] 警告：文件已保存但设备号校验不一致，请人工核对：{target}")
        return "设备号不一致"
    log(f"[{serial}] 导出失败（文件未落盘）")
    return "保存失败"


def log_row(base: Path, serial: str, target: Path, status: str):
    csv_path = base / "采集日志.csv"
    new = not csv_path.exists()
    try:
        with open(csv_path, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            if new:
                w.writerow(["时间", "设备号", "来源", "目标文件", "状态"])
            w.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        serial, "Comsoft存档", str(target), status])
    except OSError as e:
        log(f"  [警告] 写日志失败：{e}")


# ============================================================
#  主流程
# ============================================================
def _launch_comsoft(exe: str) -> bool:
    """启动 Comsoft（多路尝试，优先用开始菜单快捷方式）。

    实测（2026-10-08）：直接 `subprocess.Popen([exe])` 在非交互式会话里
    会以 0xC0000135（DLL 未找到）立即退出；而经 `cmd /c start "<快捷方式>"`
    启动可正常常驻。故顺序为：
      1) 开始菜单/桌面的 Comfort Software.lnk（最可靠）
      2) cmd /c start "<exe>"
      3) 直接 Popen（最后兜底）
    """
    exe_dir = str(Path(exe).parent)

    # 1) 快捷方式（多个候选位置）
    lnks = [
        Path(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Testo\Comfort Software.lnk"),
        Path.home() / "Desktop" / "Comfort Software.lnk",
        Path(r"C:\Users\Public\Desktop\Comfort Software.lnk"),
    ]
    for lnk in lnks:
        if lnk.exists():
            try:
                subprocess.Popen(f'start "" "{lnk}"', shell=True, cwd=exe_dir)
                return True
            except Exception:
                continue

    # 2) cmd start 直接拉起 exe
    try:
        subprocess.Popen(f'start "" "{exe}"', shell=True, cwd=exe_dir)
        return True
    except Exception:
        pass

    # 3) 兜底
    try:
        subprocess.Popen([exe], cwd=exe_dir)
        return True
    except Exception as e:
        log(f"[错误] 启动 Comsoft 失败：{e}")
        return False


def ensure_comsoft(cfg: dict, cs: Comsoft) -> bool:
    if cs.running_pid():
        return cs.connect()
    exe = cfg.get("comsoft_exe") or ""
    if not exe or not Path(exe).exists():
        log("[错误] Comsoft (cc4.exe) 未运行，且 config.json 的 comsoft_exe 路径无效")
        return False
    log(f"Comsoft 未运行，正在启动：{exe}")
    _launch_comsoft(exe)
    # 轮询等待主窗口出现（通常 3~8 秒），最多 25 秒
    deadline = time.time() + 25.0
    while time.time() < deadline:
        if cs.connect(timeout=1.5):
            return True
        time.sleep(0.4)
    return False


def run_once(base: Path, cfg: dict, force: bool) -> int:
    cs = Comsoft()
    if not ensure_comsoft(cfg, cs):
        log("[错误] 未找到 Comsoft 主窗口（确认已安装并至少运行过一次）")
        return 1
    cs.foreground()

    # 先清空所有文档标签：这是确定性状态机的前提，也避免上一轮残留干扰
    left = cs.close_tabs()
    if left:
        log(f"[提示] 仍有 {left} 个文档标签未关闭（不影响导出）")

    day = base / datetime.now().strftime(cfg["date_format"])
    day.mkdir(parents=True, exist_ok=True)

    devs = cs.devices()
    if not devs:
        log("存档树中没有设备节点（请确认设备已插入且 Comsoft 已识别）")
        return 2
    log(f"识别到设备：{sorted(devs)}")

    results = []
    for serial in sorted(devs):
        target = day / f"{serial}.vi2"
        status = export_one(cs, serial, devs[serial], target, force)
        log_row(base, serial, target,
                {"OK": "已保存(.vi2)", "跳过": "已存在，跳过"}.get(status, f"失败:{status}"))
        results.append((serial, status))

    if cfg.get("close_tabs_after_export", True):
        left = cs.close_tabs()
        if left:
            log(f"[提示] 仍有 {left} 个文档标签未关闭（不影响已导出文件）")

    ok_n = sum(1 for _, st in results if st in ("OK", "跳过"))
    bad = [(s, st) for s, st in results if st not in ("OK", "跳过")]
    log(f"\n汇总：成功 {ok_n} / {len(results)}，失败 {len(bad)}")
    for s, st in bad:
        log(f"  [{s}] {st}")
    return 2 if bad else 0


def run_watch(base: Path, cfg: dict, force: bool) -> int:
    cs = Comsoft()
    if not ensure_comsoft(cfg, cs):
        log("[错误] 未找到 Comsoft 主窗口")
        return 1
    cs.foreground()
    log(f"监视中（每 {cfg['poll_interval']:g} 秒检查一次，Ctrl+C 退出）。输出：{base}")
    try:
        while True:
            try:
                day = base / datetime.now().strftime(cfg["date_format"])
                day.mkdir(parents=True, exist_ok=True)
                for serial, dev in sorted(cs.devices().items()):
                    target = day / f"{serial}.vi2"
                    if target.exists() and not force and serial_in_file(target, serial):
                        continue
                    log(f"\n[{datetime.now():%H:%M:%S}] 发现设备 {serial}，开始导出...")
                    cs.foreground()
                    status = export_one(cs, serial, dev, target, force)
                    log_row(base, serial, target,
                            {"OK": "已保存(.vi2)"}.get(status, f"失败:{status}"))
            except Exception as e:
                log(f"[警告] 轮询异常：{e}")
            time.sleep(float(cfg["poll_interval"]))
    except KeyboardInterrupt:
        log("\n已停止监视。")
    return 0


def selftest(cfg: dict) -> int:
    log("=== 环境自检 ===")
    ok = True

    log("1) comtypes / UIA ...", )
    if check_deps() is None:
        log("   [失败] 缺少 comtypes。请先运行：pip install comtypes")
        return 1
    log("   [通过]")

    log("2) Comsoft 进程 ...")
    cs = Comsoft()
    pid = cs.running_pid()
    if not pid:
        log("   [失败] 未检测到 cc4.exe，请先打开 Testo Comfort Software")
        return 1
    log(f"   [通过] pid={pid}")

    log("3) UIA 连接（关键：不应卡死）...")
    t0 = time.time()
    if not cs.connect(timeout=20):
        log("   [失败] 无法定位 Comsoft 主窗口")
        return 1
    log(f"   [通过] 主窗口 = {cs.uia.name(cs.main)!r}"
        f"（{time.time()-t0:.1f}s）")

    log("4) 遍历控件树（关键：不应卡死）...")
    t0 = time.time()
    n = sum(1 for _ in cs.uia.walk(cs.main))
    log(f"   [通过] {n} 个元素，用时 {time.time()-t0:.1f}s")

    log("5) 存档树设备 ...")
    devs = cs.devices()
    log(f"   {'[通过]' if devs else '[提示]'} 设备：{sorted(devs) if devs else '无（请插入温度计）'}")

    log("6) 输出目录 ...")
    raw = cfg.get("output_dir") or ""
    base = Path(os.path.expandvars(raw)).expanduser() if raw else TOOL_DIR / "Reports"
    try:
        base.mkdir(parents=True, exist_ok=True)
        log(f"   [通过] {base}")
    except OSError as e:
        log(f"   [失败] {e}")
        ok = False

    log("\n自检完成。" + ("全部通过，可以正常使用。" if ok else "存在失败项，请按提示处理。"))
    return 0 if ok else 1


def main():
    cfg = load_config()
    ap = argparse.ArgumentParser(description="Testo 184 .vi2 自动导出（驱动 Comsoft）")
    ap.add_argument("--base", default=None, help="输出根目录（覆盖 config.json）")
    ap.add_argument("--comsoft", default=None, help="cc4.exe 路径（覆盖 config.json 的 comsoft_exe）")
    ap.add_argument("--watch", action="store_true", help="持续监视")
    ap.add_argument("--force", action="store_true", help="已存在也重新导出")
    ap.add_argument("--selftest", action="store_true", help="只做环境自检，不动 Comsoft")
    args = ap.parse_args()

    if check_deps() is None:
        log("[错误] 缺少 comtypes 库（UIA 自动化必需）。")
        log("       请在本工具目录执行：pip install comtypes pywin32")
        return 1

    if args.selftest:
        return selftest(cfg)

    if args.comsoft:
        cfg["comsoft_exe"] = os.path.expandvars(args.comsoft)
    raw = args.base or cfg.get("output_dir") or ""
    base = Path(os.path.expandvars(raw)).expanduser() if raw else TOOL_DIR / "Reports"
    try:
        base.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        log(f"[错误] 输出目录不可用：{base}（{e}）")
        return 1

    return run_watch(base, cfg, args.force) if args.watch else run_once(base, cfg, args.force)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        log("\n已中断。")
        sys.exit(130)
