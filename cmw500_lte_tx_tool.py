#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""CMW500 LTE 射频指标测试工具（信令模式，网线 / Raw Socket SCPI 控制）

界面: 从上到下依次为 CMW500 配置行（右侧含 AT 控制：串口自动识别/手动 debug
      指令）、电源配置行（行尾为 初始化仪表/开始测量/打开配置/打开频段表 按钮）、
      频段/带宽勾选、指标勾选，
      一个“初始化仪表”按钮 + 一个“开始测量”按钮，测完自动生成 Excel + TXT。

频段/带宽: 界面显示 B<n> / <数值>MHz；频段清单、各频段支持带宽、低中高信道
      全部取自本目录 specs\LTE\bands.json（RF Band Tool / 3GPP R19 的本地快照）。
      工具运行与刷新只读本目录，不依赖任何外部路径。

依据 R&S CMW500 Operating Manual 2.0.10:
  - 6.1 Remote Control (PDF p.139) / 31.1.2.2 Direct Socket (PDF p.2671): LAN socket, Data Port 5025
  - 21.5.2 Signal Routing (PDF p.1522): ROUT:LTE:MEAS:SCENario:CSPath (Combined Signal Path)
  - 21.5.3 Multi Evaluation Commands (PDF p.1528): 功率/EVM/ACLR/频偏 FETCh
  - 23 LTE Signaling (PDF p.1677~1776): 小区配置/连接控制/BLER 灵敏度测量
      23.4.1 General Configuration (p.1728) / 23.4.2 BLER Tests (p.1734)
      23.5.2 Connection Control (p.1737) / 23.5.10 BLER Commands (p.1768)
  - E36300 Series Programming Guide (E36311-90008, Ed.12 2025):
      p.17 通道 / p.69 MEASure:CURRent? / p.5 SOUR:VOLT + OUTP
  - E36300 Series User's Guide (E36311-90001): p.46 Using sockets, port 5025

运行:
  python cmw500_lte_tx_tool.py            # 打开 GUI
  python cmw500_lte_tx_tool.py --selftest # 无仪表自检（解析/命令构造/导出）
  python cmw500_lte_tx_tool.py --demo     # 无仪表演示测量 + 生成示例报告
"""

import argparse
import datetime
import json
import math
import os
import queue
import re
import socket
import sys
import threading
import time
import traceback
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

try:
    import serial                    # pyserial：AT 控制（USB 串口）
    import serial.tools.list_ports
except ImportError:
    serial = None

if getattr(sys, "frozen", False):
    # PyInstaller 打包后：数据目录（config/logs/result/specs）定位到 exe 所在目录，
    # 保证 exe 拷到其它电脑也能在本目录读写（与脚本版"只读本目录"策略一致）
    APP_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))


def _ensure_specs_dir():
    """冻结版首启自释放：exe 旁没有 specs/ 时，把打包内嵌的 specs 解出来。"""
    if not getattr(sys, "frozen", False):
        return
    dst_root = os.path.join(APP_DIR, "specs", "LTE")
    src_root = os.path.join(getattr(sys, "_MEIPASS", APP_DIR), "specs", "LTE")
    try:
        if os.path.isfile(os.path.join(dst_root, "bands.json")):
            return
        os.makedirs(dst_root, exist_ok=True)
        import shutil
        for name in ("bands.json", "ul_cfg.json", "refsens.json", "lte_channels.json"):
            src = os.path.join(src_root, name)
            if os.path.isfile(src):
                shutil.copy2(src, os.path.join(dst_root, name))
    except Exception:
        pass


_ensure_specs_dir()

# ----------------------------------------------------------------------------
# 运行日志
# ----------------------------------------------------------------------------


class RunLogger:
    """每次运行生成独立日志文件，写入脚本目录 logs/ 下；线程安全、行缓冲即时落盘。"""

    def __init__(self, base_dir=None, prefix="cmw500"):
        self.dir = os.path.join(base_dir or APP_DIR, "logs")
        os.makedirs(self.dir, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(self.dir, "%s_%s.log" % (prefix, stamp))
        n = 1
        while os.path.exists(path):
            path = os.path.join(self.dir, "%s_%s_%02d.log" % (prefix, stamp, n))
            n += 1
        self.path = path
        self._lock = threading.Lock()
        self._f = open(path, "w", encoding="utf-8", buffering=1)
        self._f.write("启动时间: %s\n" % datetime.datetime.now().isoformat(timespec="seconds"))

    def _stamp(self):
        return datetime.datetime.now().strftime("%H:%M:%S")

    def _raw(self, text):
        with self._lock:
            try:
                self._f.write(text)
            except Exception:
                pass

    def write(self, text, tag=None):
        line = "[%s] %s%s\n" % (self._stamp(), ("[%s] " % tag) if tag else "", text)
        self._raw(line)

    def section(self, title):
        self._raw("\n===== %s =====\n" % title)

    def kv(self, pairs):
        for k, v in pairs:
            self.write("%s: %s" % (k, v))

    def close(self):
        with self._lock:
            try:
                self._f.write("结束时间: %s\n" % datetime.datetime.now().isoformat(timespec="seconds"))
                self._f.close()
            except Exception:
                pass


# ----------------------------------------------------------------------------
# SCPI 传输（Raw Socket）
# ----------------------------------------------------------------------------


class ScpiError(Exception):
    pass


class ScpiTransport:
    def __init__(self, host, port=5025, timeout=10.0, eol="\n"):
        self.host = host
        self.port = int(port)
        self.timeout = timeout
        self.eol = eol
        self.sock = None

    @property
    def connected(self):
        return self.sock is not None

    def connect(self):
        self.close()
        s = socket.create_connection((self.host, self.port), timeout=self.timeout)
        s.settimeout(self.timeout)
        self.sock = s
        try:
            return self.idn()
        except Exception:
            return ""

    def close(self):
        if self.sock is not None:
            try:
                self.sock.close()
            except Exception:
                pass
            self.sock = None

    def write(self, cmd):
        if self.sock is None:
            raise ScpiError("未连接")
        data = (cmd.strip() + self.eol).encode("ascii", "ignore")
        self.sock.sendall(data)

    def _readline(self):
        buf = bytearray()
        while True:
            b = self.sock.recv(1)
            if not b:
                break
            if b in (b"\n", b"\r"):
                if not buf:
                    continue
                break
            buf.extend(b)
        return buf.decode("utf-8", "replace").strip()

    def _discard_inflight(self):
        """查询超时后清掉在途/残留响应，避免后续查询错位一位。

        真机 log cmw500_20261123_110012 证实：切 STYP UDCH 后 reconnect 的
        短超时(2s)查询超时，命令已发出但本端放弃读取，仪器的迟到响应残留在
        socket 中；下一次 FETC:EBlEr? 读到旧的状态串(RDY)、STAT? 读到旧的
        数值串，整体错位一位 → 大功率档所有点 tput=None 判 FAIL，结果为空。
        处理：先同步读掉在途的那一条（等它到达），再非阻塞清空可能的残留。
        """
        if self.sock is None:
            return
        # 同步读掉在途迟到响应（超时后仪器通常很快回，给略长于一次的等待）
        try:
            self.sock.settimeout(max(3.0, self.timeout))
            self._readline()
        except Exception:
            pass
        # 再非阻塞排空可能的多余残留
        try:
            self.sock.setblocking(False)
            while self.sock.recv(65536):
                pass
        except Exception:
            pass
        finally:
            try:
                self.sock.setblocking(True)
            except Exception:
                pass

    def query(self, cmd, timeout=None, expect=None):
        """发送查询并读一行响应。

        expect: 响应正则（re.search）。真实事故（cmw500_20261124_180539）：
        忙期一次查询超时后迟到响应残留，此后所有查询"读到上一条查询的响应"
        （STAT? 拿到 REL 的 CSV、PSW 拿到 ON,ADJ），且每次都能立刻读到不再
        超时 → _discard_inflight 永不触发，错位自我维持到运行结束，10MHz
        三个组合的灵敏度全部废掉。因此超时防护不够，必须有内容级校验：
        响应与命令语义不符 → 排空残留流 → 重发一次。
        """
        old = self.sock.gettimeout()
        if timeout is not None:
            self.sock.settimeout(timeout)
        try:
            self.write(cmd)
            resp = self._readline()
            if expect is not None and resp is not None \
                    and not re.search(expect, resp):
                # 响应格式与命令不符 → 流错位，排空后重发一次
                try:
                    self._discard_inflight()
                except Exception:
                    pass
                self.write(cmd)
                resp = self._readline()
            return resp
        except socket.timeout:
            # 本端超时但命令已发出：迟到响应会残留，必须清掉，否则后续查询
            # 全部错位一位（FETCH 拿到状态串、STAT 拿到数值串）。
            self._discard_inflight()
            raise
        finally:
            self.sock.settimeout(old)

    def idn(self):
        return self.query("*IDN?")


# ----------------------------------------------------------------------------
# 数值解析
# ----------------------------------------------------------------------------


def parse_values(resp):
    """解析逗号分隔的 SCPI 数值，遇到 NCAP/NAN/INF/空串等返回 None。"""
    out = []
    for tok in (resp or "").split(","):
        t = tok.strip()
        if not t:
            continue
        if re.match(r"^(NCAP|NAV|INV|INVALID|NAN|INF|\+INF|-INF|NA|9\.91E\+37)$", t, re.I):
            out.append(None)
            continue
        try:
            out.append(float(t))
        except ValueError:
            out.append(None)
    return out


# 手册 23.5.1.2 (PDF p.1736): FETCh...?/READ...?/CALCulate...? 返回数组的
# 第一个值是 "Reliability Indicator"。日志里 MOD:AVER? -> "4,INV,INV..."
# 的首位 4 就是它，不是测量值。含义：
#   0 OK | 1 Measurement Timeout | 3 Overdriven(输入过强)
#   4 Underdriven(输入电平过低) | 6 Trigger Timeout(无触发)
#   7 Acquisition Error(无法解码) | 8 Sync Error(无法同步)
#   15 Reference Frequency Error
RELIABILITY_TEXT = {
    0: "OK（测量值有效，无错误）",
    1: "Measurement Timeout（部分结果 INV / 未完成统计计数）",
    3: "Overdriven（输入信号电平过高，精度可能受损）",
    4: "Underdriven（输入上行信号电平过低，精度可能受损/测不出）",
    6: "Trigger Timeout（未检测到触发事件，测量未启动）",
    7: "Acquisition Error（无法正确解码 RF 输入信号）",
    8: "Sync Error（无法同步到 RF 输入信号）",
    15: "Reference Frequency Error（外部参考未锁定）",
}


def parse_reliability(resp):
    """取 FETCh 响应首位 Reliability 码，返回 (code, text)；无法解析返回 (None, text)。"""
    tok = (resp or "").split(",")[0].strip()
    try:
        code = int(float(tok))
    except (TypeError, ValueError):
        return None, "无法解析 Reliability（首段=%r）" % tok
    return code, RELIABILITY_TEXT.get(code, "Reliability=%d（未定义码）" % code)


def diagnose_reliability(code, cfg=None):
    """按 Reliability 码给针对性排查建议（手册 p.1680 排查清单）。

    Underdriven(4) 是本次实测日志的码：CMW 收到的 UE 上行信号太弱或没有，
    重点是"UE 是否真在 PUSCH 上行发射"和"期望功率/衰减/功控设置"，而非注册。
    """
    enp = (cfg or {}).get("enp", "23")
    rsrp = (cfg or {}).get("rsrp", "-80")
    if code == 4:
        return (
            "【输入上行信号电平过低 Underdriven】CMW 没收到足够强的 UE 上行信号。"
            "按手册 p.1680 逐项排查：\n"
            "  ① UE 是否真正进入 RRC 连接态并被调度在 PUSCH 持续上行发射"
            "（切换频段/带宽/信道后连接会进入 Signaling in Progress 重同步，"
            "刚注册完立即测可能还没开始发射，电流是否升到发射态可佐证）；\n"
            "  ② 期望标称功率 RFS:ENPower 当前=%s dBm，若远高于 UE 实际发射功率，"
            "CMW 输入衰减/量程会失配而判 Underdriven，建议按真实链路核对；\n"
            "  ③ 射频连线、外部衰减 EATT、RF 口是否与 DUT 实际连接一致；\n"
            "  ④ Uplink Power Control：PUSCh:LEVel / UL:PMAX / TPC MAXPower "
            "是否让 UL 信号强到能被 CMW 接收解码；\n"
            "  ⑤ 下行 RS EPRE 当前=%s dBm 是否足够 UE 稳定驻留。" % (enp, rsrp))
    if code == 3:
        return ("【输入信号过强 Overdriven】降低期望标称功率 ENPower 或"
                "增加外部衰减 EATT，避免 CMW 前端过载。")
    if code == 6:
        return ("【无触发 Trigger Timeout】测量未启动：确认 Combined Signal Path"
                "(CSPath) 场景已选中、触发源与信令应用一致（手册 p.1680 "
                "'verify that the combined signal path scenario is selected'）。")
    if code in (7, 8):
        return ("【无法解码/同步上行信号】UE 上行信号的频率、带宽、定时是否与"
                "小区配置一致；切换信道后 UE 是否已重新同步并发射。")
    if code == 1:
        return ("【测量超时】统计计数未完成：可增大测量超时或检查 UE 是否"
                "在整段测量窗口内持续发射。")
    return ("CMW 上行测量未返回有效结果。请确认 UE 已注册并被调度上行发射、"
            "CSPath 场景、射频连线与期望功率设置。")


def at(values, idx):
    return values[idx] if values is not None and idx < len(values) else None


def fmt(value, digits=2):
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    try:
        return ("%." + str(digits) + "f") % value
    except (TypeError, ValueError):
        return str(value)


def parse_scalar(resp):
    for v in parse_values(resp):
        if v is not None:
            return v
    return None


# ----------------------------------------------------------------------------
# 常量
# ----------------------------------------------------------------------------

# 5 项可勾选测试指标（信令模式），按 TX / RX 两个模块分组展示
# 说明：电流/电压/功率由电源默认记录，不再作为可勾选项；最大功率标签简化为 Pmax。
TX_METRIC_KEYS = ["maxpower", "aclr", "evm", "freqerr"]   # TX 指标：UE 满功率下测上行
RX_METRIC_KEYS = ["sensitivity"]                          # RX 指标：灵敏度探底
METRIC_KEYS = TX_METRIC_KEYS + RX_METRIC_KEYS
METRIC_LABELS = {
    "maxpower": "Power",
    "aclr": "ACLR",
    "evm": "EVM",
    "freqerr": "频率误差",
    "sensitivity": "灵敏度(探底)",
}

# 灵敏度功率档（RX 模块）：UE 上行发射功率工况，影响发射机自干扰下的接收灵敏度(desense)
#   high = 大功率（TPC MAXPower，UE 升到最大功率，约 23dBm）
#   low  = 小功率（TPC CLOop + CLTPower 功率，手册示例 -10dBm）
#   both = 大+小功率都测，每信道出两个灵敏度值（结果表两列）
SENS_POWER_MODES = ["high", "low", "both"]
SENS_POWER_LABELS = {
    "high": "Pmax",
    "low": "Pmin",
    "both": "Pmax+Pmin",
}

BW_CHOICES = ["B014", "B030", "B050", "B100", "B150", "B200"]
BW_TO_MHZ = {"B014": 1.4, "B030": 3.0, "B050": 5.0,
             "B100": 10.0, "B150": 15.0, "B200": 20.0}
MHZ_TO_BW = {m: c for c, m in BW_TO_MHZ.items()}

# 勾选频段时的默认带宽：只勾 10MHz（用户指定），不再自动全选该频段所有支持的带宽。
# 不支持 10MHz 的频段由 default_bw_for_band() 取最接近 10MHz 的支持带宽。
DEFAULT_BW = "B100"          # 10MHz

# 界面显示名与 CMW500 SCPI 内部键分离：
#   内部键（config.json / SCPI 命令 / 信道表 key）: 频段 "OB<n>"、带宽 "B0xx"
#     —— CMW500 的 CONF:LTE:SIGN:BAND / :CELL:BAND:DL 只接受这两种写法，不能改。
#   界面显示（复选框 / 结果表 / Excel / TXT / 日志）: 频段 "B<n>"、带宽 "1.4MHz"…


def band_label(band):
    """内部键 -> 界面频段显示名: OB1 -> B1, OB111 -> B111。"""
    s = str(band or "")
    if not s:
        return ""
    return "B" + s[2:] if s.upper().startswith("OB") else s


def bw_label(bw):
    """内部键 -> 界面带宽显示名: B014 -> 1.4MHz, B100 -> 10MHz。"""
    s = str(bw or "")
    if not s:
        return ""
    m = BW_TO_MHZ.get(s)
    return ("%gMHz" % m) if m is not None else s


def bw_code(mhz):
    """带宽(MHz) -> CMW500 带宽代码: 1.4 -> B014, 10 -> B100。"""
    m = float(mhz)
    return MHZ_TO_BW.get(m) or ("B%03d" % int(round(m * 10)))


def combo_display(band, bw):
    """组合前缀显示名，如 'B1 / 10MHz'。"""
    return "%s / %s" % (band_label(band), bw_label(bw))


# ----------------------------------------------------------------------------
# LTE 频段 / 支持带宽 / 低中高信道
#   数据源: 本目录 specs\LTE\bands.json（RF Band Tool / 3GPP R19 快照，已随工具复制）
#   算法  : TS 36.101 5.7.3，信道中心须离频段边界 >= B/2（本文件 _hml_from_range）
#           lo = max(n_offs + B*5, n_min), hi = min(n_max - B*5 + 1, n_max),
#           mid = (lo + hi)//2
#   各频段支持的带宽直接取 bands.json 的 bandwidths_mhz，不再手工维护清单。
#   加载优先级: ① 本目录 specs\LTE\bands.json 现场计算
#               ② 同目录 lte_channels.json 快照 ③ 内联兜底表
#   全程只访问本工具目录，不依赖任何外部路径。
#   重要  : 信道中心须离频段边界 >= B/2，不能取 EARFCN 范围两端。
#           例如 B1 / 10MHz 的低信道是 CH50 而非 CH0；CH0 会让整条 10MHz 信道
#           有一半越出频段 → CMW 收不到上行(Reliability=4)，UE 从注册态掉到 idle。
# ----------------------------------------------------------------------------
LOCAL_BANDS_JSON = os.path.join(APP_DIR, "specs", "LTE", "bands.json")
# UL 参考信道 RB 配置（REFSENS，3GPP TS 36.521-1 表 7.3.5-2）
#   源自 E:\LTE&NR\rf_tool\specs\LTE\ul_cfg.json，已复制到本目录 specs\LTE\，
#   格式 {band: {bw MHz: "LCRB@RBstart"}}，与 bands.json 同一份 R19 快照。
#   灵敏度测试按当前 (band,bw) 取该配置，经 UDCH 精确下发 RB 数+起始 RB。
LOCAL_UL_CFG_JSON = os.path.join(APP_DIR, "specs", "LTE", "ul_cfg.json")
# 3GPP 参考灵敏度（Full Cell BW Pow, dBm）：{band: {bw MHz: dBm}}。
#   用于灵敏度首信道"理论种子"起步，省去 -85 起的强信号空扫；经真机验证
#   B1/10MHz 理论 -96.3 vs 实测 -97.2dBm，误差 <1dB。
LOCAL_REFSENS_JSON = os.path.join(APP_DIR, "specs", "LTE", "refsens.json")
LTE_CHANNELS_JSON = os.path.join(APP_DIR, "specs", "LTE", "lte_channels.json")
CHANNEL_POSITIONS = ("低", "中", "高")

# 低/中/高 DL EARFCN，按 (频段, 带宽) 组织；元组 = (低, 中, 高)
# 仅列出该频段实际支持的带宽（TS 36.101 Table 5.6.1-1），不支持的组合会被跳过。
LTE_CHANNEL_TABLE = {
    "OB1": {"B050": (25, 300, 575), "B100": (50, 300, 550), "B150": (75, 300, 525), "B200": (100, 300, 500)},
    "OB2": {"B014": (607, 900, 1193), "B030": (615, 900, 1185), "B050": (625, 900, 1175), "B100": (650, 900, 1150), "B150": (675, 900, 1125), "B200": (700, 900, 1100)},
    "OB3": {"B014": (1207, 1575, 1943), "B030": (1215, 1575, 1935), "B050": (1225, 1575, 1925), "B100": (1250, 1575, 1900), "B150": (1275, 1575, 1875), "B200": (1300, 1575, 1850)},
    "OB4": {"B014": (1957, 2175, 2393), "B030": (1965, 2175, 2385), "B050": (1975, 2175, 2375), "B100": (2000, 2175, 2350), "B150": (2025, 2175, 2325), "B200": (2050, 2175, 2300)},
    "OB5": {"B014": (2407, 2525, 2643), "B030": (2415, 2525, 2635), "B050": (2425, 2525, 2625), "B100": (2450, 2525, 2600)},
    "OB6": {"B050": (2675, 2700, 2725), "B100": (2700, 2700, 2700)},
    "OB7": {"B050": (2775, 3100, 3425), "B100": (2800, 3100, 3400), "B150": (2825, 3100, 3375), "B200": (2850, 3100, 3350)},
    "OB8": {"B014": (3457, 3625, 3793), "B030": (3465, 3625, 3785), "B050": (3475, 3625, 3775), "B100": (3500, 3625, 3750)},
    "OB9": {"B050": (3825, 3975, 4125), "B100": (3850, 3975, 4100), "B150": (3875, 3975, 4075), "B200": (3900, 3975, 4050)},
    "OB10": {"B050": (4175, 4450, 4725), "B100": (4200, 4450, 4700), "B150": (4225, 4450, 4675), "B200": (4250, 4450, 4650)},
    "OB11": {"B050": (4775, 4850, 4925), "B100": (4800, 4850, 4900)},
    "OB12": {"B014": (5017, 5095, 5173), "B030": (5025, 5095, 5165), "B050": (5035, 5095, 5155), "B100": (5060, 5095, 5130)},
    "OB13": {"B050": (5205, 5230, 5255), "B100": (5230, 5230, 5230)},
    "OB14": {"B050": (5305, 5330, 5355), "B100": (5330, 5330, 5330)},
    "OB17": {"B050": (5755, 5790, 5825), "B100": (5780, 5790, 5800)},
    "OB18": {"B050": (5875, 5925, 5975), "B100": (5900, 5925, 5950), "B150": (5925, 5925, 5925)},
    "OB19": {"B050": (6025, 6075, 6125), "B100": (6050, 6075, 6100), "B150": (6075, 6075, 6075)},
    "OB20": {"B050": (6175, 6300, 6425), "B100": (6200, 6300, 6400), "B150": (6225, 6300, 6375), "B200": (6250, 6300, 6350)},
    "OB21": {"B050": (6475, 6525, 6575), "B100": (6500, 6525, 6550), "B150": (6525, 6525, 6525)},
    "OB22": {"B050": (6625, 7000, 7375), "B100": (6650, 7000, 7350), "B150": (6675, 7000, 7325), "B200": (6700, 7000, 7300)},
    "OB33": {"B050": (36025, 36100, 36175), "B100": (36050, 36100, 36150), "B150": (36075, 36100, 36125), "B200": (36100, 36100, 36100)},
    "OB34": {"B050": (36225, 36275, 36325), "B100": (36250, 36275, 36300), "B150": (36275, 36275, 36275)},
    "OB35": {"B014": (36357, 36650, 36943), "B030": (36365, 36650, 36935), "B050": (36375, 36650, 36925), "B100": (36400, 36650, 36900), "B150": (36425, 36650, 36875), "B200": (36450, 36650, 36850)},
    "OB36": {"B014": (36957, 37250, 37543), "B030": (36965, 37250, 37535), "B050": (36975, 37250, 37525), "B100": (37000, 37250, 37500), "B150": (37025, 37250, 37475), "B200": (37050, 37250, 37450)},
    "OB37": {"B050": (37575, 37650, 37725), "B100": (37600, 37650, 37700), "B150": (37625, 37650, 37675), "B200": (37650, 37650, 37650)},
    "OB38": {"B050": (37775, 38000, 38225), "B100": (37800, 38000, 38200), "B150": (37825, 38000, 38175), "B200": (37850, 38000, 38150)},
    "OB39": {"B050": (38275, 38450, 38625), "B100": (38300, 38450, 38600), "B150": (38325, 38450, 38575), "B200": (38350, 38450, 38550)},
    "OB40": {"B050": (38675, 39150, 39625), "B100": (38700, 39150, 39600), "B150": (38725, 39150, 39575), "B200": (38750, 39150, 39550)},
}

# 频段 DL EARFCN 全范围（仅用于展示/兜底，不再用于选信道）
LTE_BAND_DL_RANGE = {
    "OB1": (0, 599), "OB2": (600, 1199), "OB3": (1200, 1949),
    "OB4": (1950, 2399), "OB5": (2400, 2649), "OB6": (2650, 2749),
    "OB7": (2750, 3449), "OB8": (3450, 3799), "OB9": (3800, 4149),
    "OB10": (4150, 4749), "OB11": (4750, 4949), "OB12": (5010, 5179),
    "OB13": (5180, 5279), "OB14": (5280, 5379), "OB17": (5730, 5849),
    "OB18": (5850, 5999), "OB19": (6000, 6149), "OB20": (6150, 6449),
    "OB21": (6450, 6599), "OB22": (6600, 7399),
    "OB33": (36000, 36199), "OB34": (36200, 36349), "OB35": (36350, 36949),
    "OB36": (36950, 37549), "OB37": (37550, 37749), "OB38": (37750, 38249),
    "OB39": (38250, 38649), "OB40": (38650, 39649),
}

_CHANNEL_TABLE_CACHE = None
_CHANNEL_TABLE_SOURCE = ""


def _hml_from_range(dl, bw_mhz):
    """按 TS 36.101 5.7.3 计算 DL 低/中/高 EARFCN（信道须整条落在频段内）。

    B*5 = (B/2)/0.1MHz，即信道中心须离频段边界至少半个带宽，
    保证整条信道落在工作频段内。
    """
    try:
        offs = int(round(float(bw_mhz) * 5))
        lo = max(int(dl["n_offs"]) + offs, int(dl["n_min"]))
        hi = min(int(dl["n_max"]) - offs + 1, int(dl["n_max"]))
    except (KeyError, TypeError, ValueError):
        return None
    if lo > hi:
        return None
    return (lo, (lo + hi) // 2, hi)


def _table_from_bands_json(path):
    """读本目录 data\\LTE\\bands.json，现场生成全量 频段×支持带宽 信道表。"""
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    table = {}
    for num, rec in data.items():
        dl = (rec.get("earfcn") or {}).get("dl")
        if not dl:
            continue
        chans = {}
        for mhz in (rec.get("bandwidths_mhz") or []):
            h = _hml_from_range(dl, mhz)
            if h:
                chans[bw_code(mhz)] = h
        if chans:
            table["OB%d" % int(num)] = chans
    return table


def _table_from_snapshot(path):
    """读同目录 lte_channels.json（refresh_lte_channels.py 生成的便携快照）。"""
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    table = {}
    for band, rec in (data.get("bands") or {}).items():
        chans = rec.get("channels") or {}
        table[band] = {bw: (int(h["low"]), int(h["mid"]), int(h["high"]))
                       for bw, h in chans.items()}
    return table


def _load_lte_channel_table():
    """按优先级加载信道表，返回 (table, 来源说明)。

    ① 本目录 data\\LTE\\bands.json：本地权威快照，改此文件即生效；
    ② lte_channels.json：便携全量快照（bands.json 缺失/损坏时回退）；
    ③ 内联兜底表：前两者都缺失/损坏时，保证单 .py 也能独立运行。
    """
    sources = (
        (LOCAL_BANDS_JSON, lambda: _table_from_bands_json(LOCAL_BANDS_JSON)),
        (LTE_CHANNELS_JSON, lambda: _table_from_snapshot(LTE_CHANNELS_JSON)),
        ("内联兜底表", lambda: dict(LTE_CHANNEL_TABLE)),
    )
    for name, loader in sources:
        try:
            t = loader()
            if t:
                return t, name
        except Exception:
            continue
    return dict(LTE_CHANNEL_TABLE), "内联兜底表"


def channel_table():
    global _CHANNEL_TABLE_CACHE, _CHANNEL_TABLE_SOURCE
    if _CHANNEL_TABLE_CACHE is None:
        _CHANNEL_TABLE_CACHE, _CHANNEL_TABLE_SOURCE = _load_lte_channel_table()
    return _CHANNEL_TABLE_CACHE


def channel_table_source():
    """返回当前生效的信道数据来源（写进运行日志，便于核对）。"""
    channel_table()
    return _CHANNEL_TABLE_SOURCE


# ---------------------------------------------------------------- TDD 支持 ---
# 依据 R&S CMW-KM5xx/-KS5xx User Manual 1173.9628.02：
#   CONFigure:LTE:SIGN[:PCC]:DMODe FDD|TDD   （p.306，FDD 需 KS500 / TDD 需 KS550）
#   CONFigure:LTE:SIGN:CELL[:PCC]:ULDL 0..6  （p.397，仅 TDD）
#   CONFigure:LTE:SIGN:CELL[:PCC]:SSUBframe 0..9（p.397，仅 TDD）
# 实测 112703.log：FDD 模式下写 TDD band → -221 Settings conflict；
# 直连探测本仪表 DMODe TDD → -220 access permission（无 KS550 授权）。
def _bands_json():
    """读取 specs/LTE/bands.json（含每个频段的 duplex 字段），带缓存。"""
    global _BANDS_JSON_CACHE
    try:
        return _BANDS_JSON_CACHE
    except NameError:
        pass
    _BANDS_JSON_CACHE = {}
    try:
        with open(LOCAL_BANDS_JSON, "r", encoding="utf-8") as f:
            _BANDS_JSON_CACHE = json.load(f)
    except Exception:
        pass
    return _BANDS_JSON_CACHE


def band_duplex(band):
    """返回频段的双工模式 'FDD' / 'TDD'；未知按 FDD（保守）。"""
    n = str(_band_num(band))
    rec = _bands_json().get(n) or {}
    return "TDD" if str(rec.get("duplex", "")).upper() == "TDD" else "FDD"


def is_tdd_band(band):
    return band_duplex(band) == "TDD"


def tdd_cell_commands(cfg, band):
    """TDD 小区所需帧结构配置（DMODe + ULDL + SSUBframe），FDD 返回 []。

    必须在 CELL OFF 之后、BAND/CHAN 之前下发（duplex 决定 BAND 允许范围，
    手册 p.306: TDD band 仅 OB33~OB53）。ULDL/SSUB 为 3GPP TS 36.211 子帧配置，
    由 config 的 tdd_uldl / tdd_ssub 指定（默认 1 / 7，常用测例配置）。
    """
    if not is_tdd_band(band):
        return []
    p = "LTE:SIGN"

    def _ival(key, default, lo, hi):
        try:
            v = int(float((cfg or {}).get(key, default)))
        except (TypeError, ValueError):
            v = default
        return max(lo, min(hi, v))

    uldl = _ival("tdd_uldl", 1, 0, 6)
    ssub = _ival("tdd_ssub", 7, 0, 9)
    return [
        "CONF:%s:DMODe TDD" % p,
        "CONF:%s:CELL:ULDL %d" % (p, uldl),
        "CONF:%s:CELL:SSUBframe %d" % (p, ssub),
    ]


def fdd_cell_commands():
    """从 TDD 切回 FDD 时必须显式复位 duplex（否则 BAND OB1~32 被拒）。"""
    return ["CONF:LTE:SIGN:DMODe FDD"]


def detect_tdd_capability(scpi, log=None):
    """探测仪表是否支持 TDD 信令。返回 (bool, detail)。

    关键（真机实测 2026-11-25）：**CELL ON 状态下改 duplex 会被 -220
    "access permission" 拒绝**——这是状态保护，不是缺选件。探测必须先把
    小区关掉，切 TDD 验证后恢复原 duplex 与原 CELL 状态，全程不留残留。
    """
    p = "LTE:SIGN"
    def _q(cmd):
        return str(scpi.query(cmd, timeout=2.0)).strip()
    try:
        # 必须先清空错误队列：参考工具 2026-11-25-13-04-19_1756.log 实证本仪表
        # CELL OFF 后 CONF:LTE:SIGN:DMOD TDD 返回 0,"No error"。此前探测读到的
        # -220 "access permission" 是队列里此前 CELL ON 状态误操作的历史残留，
        # 不清队列会把支持 TDD 的仪表误判为不支持并剔除全部 TDD 组合。
        scpi.write("*CLS")
        orig_duplex = _q("CONF:%s:DMODe?" % p).upper()
        orig_cell = _q("SOUR:%s:CELL:STAT?" % p).upper()
    except Exception as e:
        return False, "读取 duplex/cell 状态失败: %s" % e
    was_on = not orig_cell.startswith("OFF")
    if orig_duplex.startswith("TDD"):
        return True, "当前 duplex=TDD（已启用）"
    if was_on:
        scpi.write("SOUR:%s:CELL:STAT OFF" % p)
        time.sleep(0.3)
        # 与参考工具一致：关小区后查询一次状态确认已释放，再改 duplex
        try:
            _q("SOUR:%s:CELL:STAT:ALL?" % p)
        except Exception:
            pass
    err = ""
    try:
        scpi.write("CONF:%s:DMODe TDD" % p)
        err = _q("SYST:ERR?")
    except Exception as e:
        err = str(e)
    finally:
        # 恢复原 duplex（探测总是从 FDD 起）与原 CELL 状态
        try:
            scpi.write("CONF:%s:DMODe FDD" % p)
        except Exception:
            pass
        if was_on:
            try:
                scpi.write("SOUR:%s:CELL:STAT ON" % p)
            except Exception:
                pass
        try:
            scpi.write("*CLS")   # 探测产生的错误不留给用户正常测试
        except Exception:
            pass
    if "-220" in err or "-221" in err or "permission" in err.lower():
        return False, "仪表拒绝 TDD duplex（需 CMW-KS550 选件）: %s" % err[:100]
    return True, "TDD duplex 可设置（小区关断后验证）"



_UL_CFG_CACHE = None


def ul_cfg_table():
    """加载 UL RB 配置 {band: {bwMHz: (lcrb, start_rb)}}，本地 specs/LTE/ul_cfg.json。

    键归一：频段 OB1/B1 → "1"；带宽 B100 → 10.0 → JSON 键 "10"。
    缺失/损坏返回空字典（调用方回退到满 RB@0，不影响其它指标）。
    """
    global _UL_CFG_CACHE
    if _UL_CFG_CACHE is not None:
        return _UL_CFG_CACHE
    out = {}
    try:
        with open(LOCAL_UL_CFG_JSON, encoding="utf-8") as f:
            raw = json.load(f)
        for band, bws in raw.items():
            row = {}
            for bw, spec in (bws or {}).items():
                m = re.match(r"^\s*(\d+)\s*@\s*(\d+)\s*$", str(spec))
                if m:
                    row[float(bw)] = (int(m.group(1)), int(m.group(2)))
            if row:
                out[_band_num(band)] = row
    except Exception:
        out = {}
    _UL_CFG_CACHE = out
    return out


def ul_rb_for(band, bw):
    """返回 (band,bw) 的 (LCRB, RBstart)，数据缺失返回 None。"""
    mhz = BW_TO_MHZ.get(bw)
    if mhz is None:
        try:
            mhz = float(bw)
        except (TypeError, ValueError):
            return None
    return (ul_cfg_table().get(_band_num(band)) or {}).get(mhz)


def ul_rb_display(band, bw):
    """返回 (band,bw) 的 UL RB 配置显示串 'LCRB@RBstart'（如 50@0），缺失返回 '—'。

    用于结果表固定组合列（信道之后）；数据来自本地 ul_cfg.json。
    """
    rb = ul_rb_for(band, bw)
    return ("%d@%d" % (rb[0], rb[1])) if rb else "—"


def ul_rb_display_from_combo(label):
    """从组合显示名（如 'B1 / 10MHz / 低(CH50)'）解析出 UL RB 配置显示串。

    label 拆成 band 显示名(B1) 与带宽显示名(10MHz)，反查 ul_cfg；
    解析不出返回 '—'。make_wide_row 只有 combo_label 时使用。
    """
    parts = [p.strip() for p in str(label or "").split("/")]
    if len(parts) < 2:
        return "—"
    band = parts[0]
    m = re.search(r"([\d.]+)\s*MHz", parts[1], re.I)
    if not m:
        return "—"
    return ul_rb_display(band, float(m.group(1)))


def ul_cfg_source():
    """返回 UL RB 数据来源说明（写进日志/报告）。"""
    return "本地 specs/LTE/ul_cfg.json (E:\\LTE&NR 快照, TS 36.521-1 表 7.3.5-2)"


def _band_num(band):
    m = re.search(r"\d+", str(band))
    return int(m.group()) if m else 0


def all_bands():
    """全部可测频段（内部键 OB<n>，按频段号数值升序）。"""
    return sorted(channel_table().keys(), key=_band_num)


# 频段清单由本目录 specs\LTE\bands.json 驱动，替代原先手写的 1~14/17~22/33~40 共 30 个
ALL_BANDS = all_bands()
FDD_BANDS = [b for b in ALL_BANDS if _band_num(b) < 33]   # 兼容旧引用
TDD_BANDS = [b for b in ALL_BANDS if _band_num(b) >= 33]


def band_supported_bws(band):
    """返回该频段支持的带宽代码集合（按 BW_CHOICES 顺序）。

    支持带宽完全取自本目录 data\\LTE\\bands.json 的 bandwidths_mhz；界面勾选带宽后，
    不支持的组合会自动跳过（见 _read_cfg）。
    """
    rec = channel_table().get(band) or {}
    return [w for w in BW_CHOICES if w in rec]


def default_bw_for_band(band):
    """勾选频段时默认只勾这一个带宽：优先 10MHz。

    少数频段不支持 10MHz（B31/B51/B54/B72/B73/B87/B88/B106），
    则取与其支持带宽中频率最接近 10MHz 的那个（实际为 5MHz 或 3MHz），
    保证「默认只勾一个」，且不会下发该频段不支持的非法带宽。
    """
    sup = band_supported_bws(band)
    if not sup:
        return None
    if DEFAULT_BW in sup:
        return DEFAULT_BW
    return min(sup, key=lambda w: (abs(BW_TO_MHZ[w] - BW_TO_MHZ[DEFAULT_BW]),
                                   -BW_TO_MHZ[w]))


def normalize_band_bws(cfg):
    """把配置规范成 {band: [bw, ...]}，每个频段只保留它支持的带宽。

    兼容旧版全局 bws：旧配置里 bands=["OB1","OB2"], bws=["B100"] 时，
    只把 B100 迁移给这两个旧频段，其余 61 个频段保持未选，
    避免旧全局带宽扩散到所有频段（那会产生大量用户没勾过的组合）。
    """
    raw = cfg.get("band_bws") or {}
    legacy_bands = set(cfg.get("bands") or [])
    legacy_bws = set(cfg.get("bws") or [])
    out = {}
    for b in all_bands():
        if b in raw:
            chosen = list(raw.get(b) or [])
        elif not raw and b in legacy_bands:
            chosen = [w for w in BW_CHOICES if w in legacy_bws]
        else:
            chosen = []
        supported = set(band_supported_bws(b))
        out[b] = [w for w in BW_CHOICES if w in chosen and w in supported]
    return out


def all_band_channels(band, bw=None):
    """返回某频段在指定带宽下【全部】合法 DL EARFCN（全信道扫描用）。

    与 _hml_from_range 同一 3GPP TS 36.101 5.7.3 口径：信道中心离频段边缘
    至少半个带宽，EARFCN 从 lo 到 hi 逐个枚举；低/中/高只是其中三点。
    返回 [(位置标签, earfcn), ...]，位置标签统一为「全」。
    """
    if bw is None:
        supported = band_supported_bws(band)
        bw = supported[0] if supported else ""
    mhz = BW_TO_MHZ.get(bw)
    if mhz is None:
        return []
    rec = _bands_json().get(str(_band_num(band))) or {}
    dl = (rec.get("earfcn") or {}).get("dl") or {}
    lo = hi = None
    try:
        offs = int(round(float(mhz) * 5))   # B/2 对应的 EARFCN 数（每信道 100kHz）
        lo = max(int(dl["n_offs"]) + offs, int(dl["n_min"]))
        hi = min(int(dl["n_max"]) - offs + 1, int(dl["n_max"]))
    except (KeyError, TypeError, ValueError):
        pass
    if lo is None or hi is None or lo > hi:
        # 兜底：用低/中/高表的最小/最大 EARFCN 作全扫窗口
        pts = band_channels(band, bw)
        if not pts:
            return []
        lo = min(int(e) for _, e in pts)
        hi = max(int(e) for _, e in pts)
    return [("全", earfcn) for earfcn in range(lo, hi + 1)]


def build_combos(bands, band_bws, full_scan=False):
    """测试项 = 每个已选频段 × 该频段自己勾选的带宽 × 信道集。

    默认信道集为低/中/高三点；full_scan=True 时按 TS 36.101 口径枚举该
    (频段,带宽) 下全部合法 DL EARFCN，用于全信道灵敏度探底。
    返回 (combos, skipped)，combos 元素为 (band, bw, 位置, DL EARFCN)。
    """
    combos, skipped = [], []
    for b in (bands or []):
        for w in (band_bws.get(b) or []):
            chans = all_band_channels(b, w) if full_scan else band_channels(b, w)
            if not chans:
                skipped.append(combo_display(b, w))
                continue
            for pos, earfcn in chans:
                combos.append((b, w, pos, earfcn))
    return combos, skipped


def band_bws_summary(bands, band_bws, sep="；", limit=None):
    """人可读的「频段: 各自带宽」明细，写进日志与报告。"""
    parts = []
    for b in (bands or []):
        ws = band_bws.get(b) or []
        if ws:
            parts.append("%s: %s" %
                         (band_label(b), ",".join(bw_label(w) for w in ws)))
    if limit and len(parts) > limit:
        parts = parts[:limit] + ["…"]
    return sep.join(parts) or "无"


def band_channels(band, bw=None):
    """返回某频段在指定带宽下的低/中/高 DL EARFCN: [(位置标签, earfcn), ...]。

    数据来自本目录 data\\LTE\\bands.json（3GPP R19 本地快照），已按带宽内缩 B/2，
    保证整条信道落在频段内。
    bw 为 None 时取该频段支持的第一个带宽（兼容旧调用）。
    该频段不支持该带宽时返回空列表（调用方应跳过此组合，不下发非法信道）。
    窄频段下低/中/高可能重合，此时去重，避免同一信道重复测量。
    """
    rec = channel_table().get(band)
    if not rec:
        return []
    if bw is None:
        supported = band_supported_bws(band)
        bw = supported[0] if supported else ""
    h = rec.get(bw)
    if not h:
        return []
    out, seen = [], set()
    for pos, earfcn in zip(CHANNEL_POSITIONS, h):
        if earfcn in seen:
            continue
        seen.add(earfcn)
        out.append((pos, earfcn))
    return out


def build_channel_cmd(cfg, earfcn):
    """构造下行信道配置命令（手册 23.4.1: RFSettings:CHANnel:DL）。"""
    return "CONF:LTE:SIGN:RFS:CHAN:DL %d" % int(earfcn)

RF_CONNECTORS = ["RF1C", "RF2C", "RF3C", "RF4C", "RFAC", "RFBC"]  # 手册 p.1524
RX_CONNECTORS = ["RX1", "RX2"]

PSU_CHANNELS = ("P6V", "P25V", "N25V")  # E36312A = CH1 / CH2 / CH3
PSU_VOLT_RANGE = {
    "P6V": (0.0, 6.18, "0 ~ 6.18 V"),
    "P25V": (0.0, 25.75, "0 ~ +25.75 V"),
    "N25V": (-25.75, 0.0, "-25.75 ~ 0 V"),
}
PSU_DEFAULT_SAMPLES = 3

# 结果目录
RESULT_DIR = os.path.join(APP_DIR, "result")
CONFIG_PATH = os.path.join(APP_DIR, "config.json")

CONFIG_DEFAULTS = {
    "cmw_ip": "192.168.1.100", "cmw_port": "5025",
    "rf": "RF1C", "rx": "RX1",
    "psu_enable": True, "psu_ip": "", "psu_port": "5025",
    "psu_channel": "P6V", "psu_voltage": "3.800", "psu_readv": True,
    "freq": "1850", "enp": "23", "eatt": "0", "umar": "3",
    # 每个频段独立选择的带宽 {band: [bw, ...]}；bws 仅作兼容/汇总字段
    "bands": [], "bws": ["B100"], "band_bws": {},
    # ASRrun V2.0.7 对比：初始建网电平 -85dBm（比原 -80dBm 更接近 UE 实际运行电平，
    # 避免强信号下 UE 自动降功率的瞬态过程）。
    "rsrp": "-85", "pmax": "23", "dl_rb": "N50", "dl_mod": "QPSK",
    # FDC 频率响应校准表（ASRrun V2.0.7 对比参考）。按频率分段补偿线缆/连接器损耗，
    # 仪器自动根据当前测试频率从表中查表插值。格式：freq_MHz:loss_dB 逗号分隔，频率严格递增。
    # 默认值为通用射频线缆的经验值（450~3600MHz 0.3~1.0dB 分段），请根据实际线缆长度/材质校准。
    # 设为空字符串或 false 可关闭 FDC。
    "freq_loss_segments": "450:0.3, 800:0.3, 1000:0.5, 1500:0.5, 2000:0.7, 2500:0.7, 3600:1.0",
    "freq_loss_enable": True,
    # 灵敏度口径：仪器设置/扫描的是 RS EPRE (dBm/15kHz)，报告换算为 Full Cell BW Pow (dBm)
    #   FullCellBW = RS_EPRE + 10log10(12*N_RB)，N_RB=6/15/25/50/75/100 (1.4/3/5/10/15/20MHz)
    # 扫描下限按 3GPP REFSENS 留余量：10MHz 参考灵敏度约 -97dBm(FullBW) => RS EPRE 约 -125，
    #   故默认 stop=-135；起点 -85 RS EPRE（避开强信号浪费扫描时间）。
    #   三级扫描：coarse 大步进 -> near 中步 -> step 精步(0.1dBm)，兼顾速度与精度。
    "sens_start": "-85", "sens_stop": "-135", "sens_step": "0.1", "sens_near": "1",
    "sens_coarse": "5",
    # 热点续扫独立参数：已有同(band,bw)上一信道实测值时用较小 margin + 较小 coarse 步长，
    # 快速逼近拐点后直接进入二分精扫，避免原 near=1dB 步进 6-8 点空扫强信号区。
    "sens_hotspot_coarse": "2",
    "sens_hotspot_margin": "3", "sens_hotspot_span": "3",
    # 理论种子起步 margin：独立参数，未显式配置时自动沿用 sens_hotspot_margin
    # （保持向后兼容——原 sens_hotspot_margin 同时服务两种起步场景）。
    # 种子值与实际灵敏度可能差 ±4dB，需较大 margin 保证覆盖但不浪费空扫。
    "sens_seed_margin": None,
    # 灵敏度判定：Throughput(%) ≥ sens_throughput(%) 为 PASS（默认 95%）。
    # 判定量取 FETC:LTE:SIGN:EBLer:REL? 的第 2 值 Total TPut(%)（idx 1），
    # 与 ASRrun 参考工具及 3GPP TS 36.521-1 的 RefSens 判据一致
    # （2026-11-25 实测核对：参考工具数据表 TPut 与响应 idx1 完全吻合；
    # 旧取 idx4 是另一个略偏乐观的量，已修正）。
    "sens_throughput": "95",
    # （遗留键，语义已废弃：REL? idx1 实为吞吐量而非 TotalBlocks 百分比，
    # 按 idx1<50% 判"统计不足"的旧机制 2026-11-25 移除，统计有效性改由
    # BLER 值是否存在判定。键保留仅为 config.json 向后兼容。）
    "sens_min_total": "50",
    "sens_min_total_low": "20",
    # UE 掉线特征门限：TPut(%) 低于该值时才触发注册态确认。注意 TPut 低
    # 本身不等于掉线（极弱信号下 UE 仍在网也收不到 PDCCH，真机 145949.log 有
    # TPut=0~0.17% 但 PSW=ATT 的点），必须再查 PSW:STAT=ON 才判掉线并停止二分。
    "sens_drop_total": "10",
    # 重附着快速升级窗口（秒）：ACT CONN 后若该时间内未进 ATT/CEST 即判定硬掉线，
    # 直接重建小区，不再死等 timeout×2。真机 144321.log：成功重附着 1 秒内 CEST，
    # 失败则 20 秒全程 ON——5 秒足够区分，可为组合 3 省约 40 秒。
    "attach_quick_timeout": "5",
    # 频段/带宽快速切换窗口（秒）：强信号下直接改 BAND/带宽/信道 + 幂等 CELL ON
    # 后，该时间内 UE 保持/恢复 ATT/CEST 即跳过完整小区重建（参考 ASRrun 工具），
    # 省约 8~10s/次；超时回退 retune_cell 重建小区，最坏情况=旧行为。
    # 快切等待窗口：真机 2026-11-24 实测 1.4M→10M 约 3.6s 恢复注册，
    # 但 10M→1.4M 方向仪表忙期更长（首次查询 ~5.8s 超时，ATT ~6.5s 才出现），
    # 5s 窗口会错过并白白回退完整重建；8s 两个方向都能命中。
    "fast_retune_timeout": "8",
    # 热切后 UE 掉线且无 AT 口时的被动等待（参考工具 log 实测 UE 自主重附着
    # ~26s，30s 覆盖抖动）；有 AT 口时不等，立即分流 AT 恢复
    "fast_retune_passive_wait": "30",
    # 首窗口未确认后的忙期宽限（2026-11-25 新增）：→10MHz 实测仪表查询黑洞
    # ~10-12s 超过首窗口 8s，超时/空响应不能证明 UE 掉线；宽限期内拿到合法
    # ATT/CEST 视为 UE 自恢复成功，合法 ON/OFF/DISC 才判掉线交 AT（避免
    # CFUN=0 打断 UE 正在进行的自主重附着）
    "fast_retune_busy_wait": "25",
    # ASRrun V2.0.7 对比参考：600 子帧 ≈ 6 秒/点，比 1000 子帧省 40% 时间，
    # 拐点处 BLER 从 0→10% 跃迁远大于 600/1000 子帧统计误差。
    "bler_subframes": "600", "bler_timeout": "10",
    # TPC MAXPower 后 UE 升功率需要稳定；真机偶发 Rel=3 Overdriven，2s→4s。
    "tpc_settle": "4", "stats": "20",
    # 切信道重同步/测量重试（手册 p.1681：切信道使连接进入 Signaling in Progress，
    # UE 需重新同步才可能被调度上行发射）。仅走配置默认值，不在界面显示。
    "chan_settle": "1.5", "mev_retry": "2", "mev_retry_gap": "1.0",
    # 切频段/带宽重附着后加 2s 上行稳定等待：真机 134654.log 组合 7
    # 切到 5MHz 后 UE 刚 ATT 就被 MAXPower + MEV，上行未启动导致全 INV、
    # Reliability=3 Overdriven 中止。等待后复查注册再交 TX 测量。
    "attach_settle": "2",
    # 初始化完成提示框自动关闭时间（毫秒）：无需人工点击确认，避免阻塞后续操作
    "init_popup_ms": "3000",
    # AT 控制串口（USB 口接 DUT 模组 AT 口）：空/"auto" = 启动时自动识别（AT 握手探测）
    "at_port": "",
    # AT 指令输入框内容：关闭工具时保存，下次启动恢复（多条用 ; 分隔依次执行）
    "at_cmd": "",
    # TDD 小区帧结构（3GPP TS 36.211）：ULDL=上下行子帧配置 0~6，SSUB=特殊子帧 0~9
    # 仅 TDD 频段（OB33~OB53）生效；需 CMW-KS550 选件，启动时自动探测授权。
    "tdd_uldl": "1",
    "tdd_ssub": "7",
    # TDD 授权探测：默认关闭（0=按仪表支持处理，不剔除 TDD 组合）。
    # 参考工具 2026-11-25-13-04-19_1756.log 实证本仪表 CELL OFF 后
    # CONF:LTE:SIGN:DMOD TDD 无错误；探测本身会切换 duplex 有副作用，
    # 仅在用户显式设 1 时才执行（此时会先 *CLS 清错误队列避免误判）。
    "tdd_detect": "0",
    # 掉线自动恢复：AT CFUN=0→1 强制重注册（1=启用，0=禁用回退纯仪表侧恢复）
    "at_auto_recover": "1",
    # CFUN=1 后等待 UE 回网的最长秒数（超时回退重建小区；ASR cat.1 实测 5~20s）
    "at_recover_wait": "60",
    "evm_lim": "17.5", "aclr_lim": "30", "freq_lim": "100", "sens_lim": "",
    # 测试场景（TX/RX 公用）：high=Pmax / low=Pmin / both=Pmax+Pmin（两档都测，双列结果）
    # 旧配置键 sens_power_mode 保留为兼容别名，load_config / _read_cfg 会双向同步。
    "power_scenario": "high",
    "sens_power_mode": "high",
    "sens_low_power": "",  # Pmin 场景的 CLOop dBm（空 → UE 最小发射功率；填值 → CLOop 功率，范围 -50~23）
    # 全信道扫描：勾选后每个 (频段,带宽) 枚举全部合法 DL EARFCN，逐信道探底灵敏度
    "full_scan": False,
    "metrics": {k: True for k in METRIC_KEYS},
}


def load_config():
    out = dict(CONFIG_DEFAULTS)
    out["metrics"] = dict(CONFIG_DEFAULTS["metrics"])
    try:
        if os.path.isfile(CONFIG_PATH):
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            for k, v in data.items():
                out[k] = v
            m = dict(CONFIG_DEFAULTS["metrics"])
            m.update(data.get("metrics") or {})
            out["metrics"] = m
    except Exception:
        pass
    # 测试场景：新键 power_scenario 与旧键 sens_power_mode 双向兼容
    if "power_scenario" in out:
        out["sens_power_mode"] = out["power_scenario"]
    else:
        out["power_scenario"] = out.get("sens_power_mode", "high")
    # 保证 metrics 覆盖全部指标键
    for k in METRIC_KEYS:
        out["metrics"].setdefault(k, True)
    return out


def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


# ----------------------------------------------------------------------------
# SCPI 命令构造（信令模式）
# ----------------------------------------------------------------------------


def build_fdc_commands(cfg):
    """构造 FDC 频率响应校准命令序列（CMW 手册 10.4.1 / ASRrun V2.0.7 对比）。

    按频率分段补偿线缆/连接器损耗，CMW 自动根据当前测试频率从表中查表插值。
    格式: "freq_MHz:loss_dB, freq_MHz:loss_dB, ..." 频率严格递增。

    返回 SCPI 列表（可能为空），不直接 query，让上层统一走 transport。
    """
    if not cfg.get("freq_loss_enable"):
        return []
    raw = str(cfg.get("freq_loss_segments", "")).strip()
    if not raw:
        return []
    # 解析 "450:0.3, 800:0.3, ..."
    segs = []
    for tok in raw.split(","):
        tok = tok.strip()
        if not tok:
            continue
        if ":" not in tok:
            continue
        fm, lm = tok.split(":", 1)
        try:
            freq_hz = float(fm) * 1e6  # MHz → Hz
            loss_db = float(lm)
        except ValueError:
            continue
        segs.append((freq_hz, loss_db))
    if len(segs) < 2:
        return []
    segs.sort(key=lambda s: s[0])  # 确保频率递增
    table = "losstable"
    rf_port = cfg.get("rf", "RF1C")
    cmds = []
    cmds.append("CONF:BASE:FDC:RCL")  # 先 DeACtivate 所有 FDC
    cmds.append("CONF:BASE:FDC:CTABle:DELete '%s'" % table)  # 删旧表（有则删）
    cmds.append("CONF:BASE:FDC:CTABle:CREATE '%s', %.0f, %.1f" % (table, segs[0][0], segs[0][1]))
    for freq_hz, loss_db in segs[1:]:
        cmds.append("CONF:BASE:FDC:CTABle:ADD '%s', %.0f, %.1f" % (table, freq_hz, loss_db))
    cmds.append("CONFigure:FDCorrection:ACTivate %s, '%s', TX" % (rf_port, table))
    cmds.append("CONFigure:FDCorrection:ACTivate %s, '%s', RX" % (rf_port, table))
    return cmds


def lookup_loss_mhz(cfg, mhz):
    """按频率从 freq_loss_segments 分段表线性插值查损耗 dB。

    分段表格式: "450:0.3, 800:0.3, 1000:0.5, ..."（freq_MHz:loss_dB，频率递增）。
    低于第一段频率 → 用第一段损耗；高于最后一段 → 用最后一段损耗；中间线性插值。
    返回 None 表示没启用或配置无效。
    """
    if not cfg.get("freq_loss_enable"):
        return None
    raw = str(cfg.get("freq_loss_segments", "")).strip()
    if not raw:
        return None
    segs = []
    for tok in raw.split(","):
        tok = tok.strip()
        if ":" not in tok:
            continue
        fm, lm = tok.split(":", 1)
        try:
            segs.append((float(fm), float(lm)))
        except ValueError:
            continue
    if len(segs) < 2:
        return None
    segs.sort(key=lambda s: s[0])
    f = float(mhz)
    if f <= segs[0][0]:
        return segs[0][1]
    if f >= segs[-1][0]:
        return segs[-1][1]
    for i in range(len(segs) - 1):
        if segs[i][0] <= f <= segs[i + 1][0]:
            f0, l0 = segs[i]
            f1, l1 = segs[i + 1]
            if f1 == f0:
                return (l0 + l1) / 2.0
            return l0 + (l1 - l0) * (f - f0) / (f1 - f0)
    return None


def earfcn_to_dl_center_mhz(band, earfcn):
    """从原始 bands.json 数据 + DL EARFCN 推算下行中心频率（MHz，TS 36.101 公式）。

    f_DL = F_DL_low + 0.1 × (N_DL - N_DL_offset)

    注意：_load_lte_channel_table() 是简化表（只有 EARFCN 列表），
    这里直接读原始 specs/LTE/bands.json 取 f_dl_low 和 earfcn.dl.n_offs。
    """
    if not band or earfcn is None:
        return None
    try:
        with open(LOCAL_BANDS_JSON, encoding="utf-8") as f:
            raw = json.load(f)
    except Exception:
        return None
    # bands.json 的 key 是数字字符串 "1", "8", "41"；传入可能是 "OB1" 格式
    band_num = str(band).lstrip("OoBb")
    b = raw.get(band_num)
    if not b:
        return None
    try:
        n_offs = b["earfcn"]["dl"]["n_offs"]
        f_dl_low = b["f_dl_low"]
    except (KeyError, TypeError):
        return None
    return f_dl_low + 0.1 * (int(earfcn) - int(n_offs))


def build_signaling_init(cfg, band="", bw=""):
    """构造 LTE 信令初始化命令序列（手册 23.4.1 General Configuration）。

    返回 (commands, meta)。小区开/连接建立命令单独返回，便于按状态轮询。
    """
    c = []
    p = "LTE:SIGN"           # 信令实例 1 可省略 <i>
    m = "LTE:MEAS"           # Multi Evaluation 实例

    def f(key):
        try:
            return float(cfg.get(key, 0))
        except (TypeError, ValueError):
            return 0.0

    # 系统复位
    c.append("*RST")
    c.append("*CLS")

    # 路由（手册 p.1729）：SCEL <out>,<rx>,<in>,<tx>
    c.append("ROUT:%s:SCEN:SCEL %s,%s,%s,TX1" % (p, cfg["rf"], cfg["rx"], cfg["rf"]))

    # FDC 频率响应校准（ASRrun V2.0.7 对比）——插在路由之后、频段配置之前
    c.extend(build_fdc_commands(cfg))

    # Duplex 模式必须先于 BAND 设置（手册 p.306：BAND 允许范围取决于 duplex；
    # FDD=OB1~32，TDD=OB33~OB53）。TDD 还需帧结构 ULDL/SSUB（p.397）。
    if is_tdd_band(band):
        c.extend(tdd_cell_commands(cfg, band))
    else:
        c.extend(fdd_cell_commands())
    # 频段 / 带宽 / 信道
    if band:
        c.append("CONF:%s:BAND %s" % (p, band))
    if bw:
        c.append("CONF:%s:CELL:BAND:DL %s" % (p, bw))
    # 期望功率模式：ULPC（闭环功率控制，对齐 ASRrun 参考工具 1756/3008.log）。
    # MANual 模式下 UE 上行不被 TPC 驱动，KEEP 型 RMC 会因缺 TBS 显式值而
    # 完全不发射 PUSCH → MEV 全 INV/NCAP（11-25 135456~141808 四次实证）。
    # 参考工具用 ULPC + TPC SET 让 UE 持续发射，TX MEV 正常出数。
    c.append("CONF:%s:RFS:ENPMode ULPC" % p)
    c.append("CONF:%s:RFS:ENPower %.1f" % (p, f("enp")))
    c.append("CONF:%s:RFS:UMARgin %.1f" % (p, f("umar")))
    c.append("CONF:%s:RFS:EATT:OUTP %.1f" % (p, f("eatt")))
    # 外接衰减器双向共用（FDD 双工同缆），UL 输入期望功率同步补偿（手册 p.1729 成对示例）
    c.append("CONF:%s:RFS:EATT:INP %.1f" % (p, f("eatt")))

    # 下行参考功率与上行最大功率（手册 p.1729~1730）
    c.append("CONF:%s:DL:RSEPre:LEVel %.1f" % (p, f("rsrp")))
    c.append("CONF:%s:UL:PMAX %.1f" % (p, f("pmax")))

    # 连接类型为 RMC，配置上下行 RMC（手册 p.1734 / p.1763）
    # UET RMC：显式把 UE 类型设为 RMC。参考工具 1756.log 每个测试阶段都发
    # （109/882 行）；KEEP 模式下缺这条会导致 UE 虽注册 ATT 却不按 RMC
    # 调度上行 → MEV 全 INV/NCAP（140933.log 实证）。
    c.append("CONF:%s:CONN:UET RMC" % p)
    c.append("CONF:%s:CONN:STYP RMC" % p)
    # TX 满 RB 必须按首个组合的带宽换算：初始化即按首组合带宽建小区，
    # 首带宽随后走快速路径（只改 BAND/BW/CHAN）不再重配 RMC，若沿用 cfg
    # 里的 dl_rb（默认 N50=10MHz）会导致 1.4MHz 实际跑 50 RB。
    # 与 retune_cell 的 ②a 逻辑一致：B014→N6，其余 bw*5。
    _init_rb = cfg["dl_rb"]
    if bw:
        _init_rb = "N%d" % (6 if bw == "B014"
                            else int(BW_TO_MHZ.get(bw, 10) * 5))
    c.append("CONF:%s:CONN:RMC:DL %s,%s,KEEP" % (p, _init_rb, cfg["dl_mod"]))
    c.append("CONF:%s:CONN:RMC:UL %s,QPSK,KEEP" % (p, _init_rb))
    # 参考工具 ASRrun 显式固定 RB 起始位置为 LOW；不写时依赖仪表默认，
    # 跨带宽/STYP 切换后可能出现位置残留，故初始化即固定。
    c.append("CONF:%s:CONN:RMC:RBP:DL LOW" % p)
    c.append("CONF:%s:CONN:RMC:RBP:UL LOW" % p)

    # Multi Evaluation 走 Combined Signal Path，master = 信令应用
    # 手册 p.1524: ROUT:LTE:MEAS:SCENario:CSPath <Master>
    # 参数用手册规定的信令应用实例名 'LTE Sig1'（避免 -222 Data out of range）。
    # TX 全 INV 与 CSPath 取值无关：140933.log 回退 'LTE:SIGN' 后 -222 重现，
    # MEV 同样 Reliability=3 全 INV；根因是 KEEP 模式下缺 UET RMC / 切信道后
    # 未重发 RMC:UL（见 build_signaling_init 与主循环补强）。
    c.append("ROUT:%s:SCEN:CSPath 'LTE Sig1'" % m)

    meta = {
        "p": p, "m": m,
        "cell_on": "SOUR:%s:CELL:STAT ON" % p,
        "cell_state_q": "SOUR:%s:CELL:STAT:ALL?" % p,
        "conn_cmd": "CALL:%s:PSW:ACT CONN" % p,
        "conn_state_q": "FETC:%s:PSW:STAT?" % p,
    }
    return c, meta


def build_tx_queries(cfg):
    """连接态下测 UE 上行的测量启动 + FETCh 查询（手册 21.5.3 / CSPath）。

    关键: *RST 后 MEV 处于 OFF，必须先 INIT 触发 single-shot，等状态 RDY
    后 FETCh 才有有效结果（手册 PDF p.1510: "After a *RST, the measurement is
    switched off. Use INIT... and retrieve the results using FETCh...?").
    """
    m = "LTE:MEAS"
    return {
        "init": "INIT:%s:MEV" % m,
        "state": "FETC:%s:MEV:STAT?" % m,
        "mod_avg": "FETC:%s:MEV:MOD:AVER?" % m,
        "mod_ext": "FETC:%s:MEV:MOD:EXTR?" % m,
        "aclr": "FETC:%s:MEV:ACLR:AVER?" % m,
    }


def build_mev_config_commands(cfg, band=""):
    """显式配置 Multi Evaluation（对齐 ASRrun 参考流程，防状态残留）。

    只开启本工具真正 FETCh 的结果：TXM（功率/EVM/频偏）与按需 ACLR；
    频谱掩模 / 杂散发射关闭，避免无关测量拖慢单次 INIT。

    MSUBframes 第三参数（子帧起始偏移）必须按双工模式区分：参考工具
    2026-11-25-13-04-19_1756.log 实证 FDD=0、TDD=2。TDD 上行子帧从
    subframe 2 起（0/5 为 DL、1 为特殊子帧），偏移 0 会把测量窗口落在
    DL 子帧 → 上行结果全 INV（112703.log 的 TDD 无效数据根因之一）。
    """
    m = "LTE:MEAS"
    metrics = cfg.get("metrics") or {}
    aclr_on = bool(metrics.get("aclr", False))
    _sub_off = 2 if is_tdd_band(band) else 0
    return [
        "CONF:%s:MEV:REP SING" % m,
        "CONF:%s:MEV:MSUBframes 0,10,%d" % (m, _sub_off),
        "CONF:%s:MEV:MOD:MSCHeme QPSK" % m,
        "CONF:%s:MEV:RES:TXM ON" % m,
        "CONF:%s:MEV:RES:ACLR %s" % (m, "ON" if aclr_on else "OFF"),
        "CONF:%s:MEV:RES:SEMask OFF" % m,
        "CONF:%s:MEV:RES:IEMissions OFF" % m,
    ]


def apply_mev_config(scpi, cfg, log=None, band=""):
    """下发显式 MEV 配置；每次 TX 测量前调用，保证配置不被切信道/STYP 污染。"""
    for cmd in build_mev_config_commands(cfg, band):
        scpi.write(cmd)
        if log:
            log("[SCPI] " + cmd)


def build_force_maxpower_cmd(cfg):
    """构造强制 UE 满功率发射的命令（手册 PDF p.1711 / p.1751 23.5.6）。

    PMAX 只是"允许的最大功率上限"，不会让 UE 主动升功率；UE 上行功率由
    开环标称功率 + TPC 决定，*RST 默认闭环功率 CLTPower=-20 dBm —— 这正是
    实测只有约 -17 dBm 的原因。CONF:LTE:SIGN:UL:TPC:SET MAXPower 通过 +1 dB
    步进把 UE 命令到最大功率（受 PMAX 与 UE 功率等级限制），连接建立后自动执行。
    """
    p = "LTE:SIGN"
    return {
        "set_max": "CONF:%s:UL:TPC:SET MAXPower" % p,
        "settle": max(0.0, float(cfg.get("tpc_settle", 4) or 4)),
    }


def force_max_power(scpi, cfg, log=None):
    """下发 TPC MAXPower 并等待 UE 升到最大功率、功率稳定后再测量。"""
    mp = build_force_maxpower_cmd(cfg)
    scpi.write(mp["set_max"])
    if log:
        log("[SCPI] " + mp["set_max"])
    settle = mp["settle"]
    if settle:
        if log:
            log("[功率] 已命令 UE 升至最大功率，等待 %.1fs 稳定后测量" % settle)
        time.sleep(settle)


def build_sens_power_cmd(cfg, mode, power_min):
    """构造灵敏度功率档的 UE 上行功率命令列表（手册 PDF p.1711/1752/1753）。

    high = 大功率：CONF:LTE:SIGN:UL:TPC:SET MAXPower（强制最大功率）
    low  = 小功率：
        power_min is None → CONF:LTE:SIGN:UL:TPC:SET MINPower（UE 最小发射功率）
        有值 → UPLink:TPC:CLTPower <P> + UL:TPC:SET CLOop（闭环功率 -50~23dBm）
    返回 (命令列表, 描述字符串)。
    """
    p = "LTE:SIGN"
    if mode == "low":
        if power_min is not None:
            cmds = ["CONF:%s:UPLink:TPC:CLTPower %.1f" % (p, power_min),
                    "CONF:%s:UL:TPC:SET CLOop" % p]
            return cmds, "CLOop 功率 %.1f dBm" % power_min
        else:
            cmds = ["CONF:%s:UL:TPC:SET MINPower" % p]
            return cmds, "MINPower（UE 最小发射功率）"
    cmds = ["CONF:%s:UL:TPC:SET MAXPower" % p]
    return cmds, "MAXPower（UE 最大发射功率）"


def set_sens_power(scpi, cfg, mode, log=None):
    """设置灵敏度测量的 UE 上行功率档，并等待稳定。"""
    # sens_low_power 输入框：空 → None（最小功率）；有值 → float（CLOop 功率）
    raw = (cfg.get("sens_low_power") or "").strip()
    power_min = None  # 默认最小功率
    if raw:
        try:
            power_min = float(raw)
            if not (-50 <= power_min <= 23):
                power_min = None  # 超出范围回退最小功率
        except ValueError:
            power_min = None
    cmds, desc = build_sens_power_cmd(cfg, mode, power_min)
    for cmd in cmds:
        scpi.write(cmd)
        if log:
            log("[SCPI] " + cmd)
    if log:
        log("[灵敏度功率] UE 上行发射功率档：%s" % desc)
    try:
        settle = max(0.0, float(cfg.get("tpc_settle", 4) or 4))
    except (TypeError, ValueError):
        settle = 4.0
    if settle:
        time.sleep(settle)


def build_udch_commands(band, bw):
    """构造灵敏度 REFSENS 参考信道的 UDCH 配置命令（手册 p.1766~1767）。

    UL: UDCHannels:UL <NumberRB>,<StartRB>,QPSK,KEEP —— 精确 LCRB@RBstart；
    DL: UDCHannels:DL <NumberRB>,<StartRB>,QPSK,KEEP —— 满 RB@0（LTE 无独立 DL 表）。
    RB 数必须是仪表接受的枚举（1/2/3/4/5/6/8/10/12/15/16/18/20/24/25/27/
    30/36/40/48/50/75/80/100）；权威表里的 LCRB 均落在该集合内。
    """
    p = "LTE:SIGN"
    dl_rb = int(BW_TO_MHZ.get(bw, 10) * 5)  # 满 RB：1.4M→6(特判)、3→15、10→50、20→100
    if bw == "B014":
        dl_rb = 6
    # 数据缺失回退到该带宽满 RB@0（满 RB 数本身是仪表合法枚举，窄带也不会越界）
    rb = ul_rb_for(band, bw) or (dl_rb, 0)
    lcrb, start = rb
    return {
        "set_udch": "CONF:%s:CONN:STYP UDCH" % p,
        "dl": "CONF:%s:CONN:UDCH:DL %d,0,QPSK,KEEP" % (p, dl_rb),
        "ul": "CONF:%s:CONN:UDCH:UL %d,%d,QPSK,KEEP" % (p, lcrb, start),
        "rb": rb, "dl_rb": dl_rb,
        "source": ul_cfg_source(),
    }


def restore_rmc_commands(cfg, bw=""):
    """灵敏度测完切回 RMC（恢复后续 TX 组合的满功率 RMC 连接）。

    必须按当前 bw 恢复满 RB：灵敏度在 1.4MHz 测完后若恢复成 cfg 的全局
    N50（10MHz），小区带宽与 RMC RB 不匹配，后续 TX 测量会异常。
    """
    p = "LTE:SIGN"
    rb = cfg["dl_rb"]
    if bw:
        rb = "N%d" % (6 if bw == "B014"
                      else int(BW_TO_MHZ.get(bw, 10) * 5))
    return [
        "CONF:%s:CONN:STYP RMC" % p,
        "CONF:%s:CONN:RMC:DL %s,%s,KEEP" % (p, rb, cfg["dl_mod"]),
        "CONF:%s:CONN:RMC:UL %s,QPSK,KEEP" % (p, rb),
        "CONF:%s:CONN:RMC:RBP:DL LOW" % p,
        "CONF:%s:CONN:RMC:RBP:UL LOW" % p,
        # 灵敏度段打开的 OCNG 必须关闭：残留 OCNG 噪声会抬高 UE 上行功率，
        # 导致后续 TX 测量 CMW 前端 Overdriven（Reliability=3）中止。
        "CONF:%s:DL:OCNG OFF" % p,
    ]


def apply_sens_rb(scpi, band, bw, cfg, log=None, recheck=None):
    """灵敏度前按 3GPP TS 36.521-1 表 7.3.5-2 配置参考信道 RB；返回命令 dict。

    表 7.3.5-2 的 LCRB@RBstart 分两类，仪表侧命令能力不同（CMW500 手册
    p.1763~1766 已核实）：

      * RBstart = 0（多数频段全部带宽）：RMC 即可精确表达 —— 
        RMC:UL N<LCRB>,QPSK,KEEP + RMC:RBP:UL LOW（LOW = 从第 0 个 RB 起）。
        这是参考工具 ASRrun 全部日志使用的路径。
      * RBstart > 0（表内共 48 项，分布在 27 个 band，典型如 B5/B6/B8/B11/
        B27/B30 的 10MHz 25@25）：**只能走 UDCH(User Def)** —— 
        UDCH:UL <LCRB>,<RBstart>,QPSK,KEEP。RMC 的 RBP:UL 枚举只有 LOW|HIGH，
        不存在 P<start> 任意起始位置（ASRrun 参考日志对 25@25 发 RBP:UL P25
        被仪表回 -203 "option missing"，实际并未生效）；MD 明确 UDCH 需 KS510
        选件。

    TX/RX 分离：RX 用本函数按表下发的 LCRB@RBstart；TX full RB 由
    build_signaling_init / retune_cell / restore_rmc_commands 的 RMC KEEP 命令
    独立负责，两条路径不共用同一条命令。OCNG ON 由 measure_sensitivity
    统一打开（§7.3.5 要求 one-sided dynamic OCNG）。

    切 STYP/RBP/RMC:UL 会使 UE 掉线重连，故下发后按 recheck 回调主动重新附着。
    """
    cmds = build_udch_commands(band, bw)
    def w(cmd):
        scpi.write(cmd)
        if log:
            log("[SCPI] " + cmd)
    lcrb, start = cmds["rb"]
    p = "LTE:SIGN"
    if start > 0:
        # RBstart 非 0：RMC 无 P<start> 能力，必须 UDCH（需 KS510）
        mode = "UDCH"
        cmds["rb_applied"] = (lcrb, start)
        if log:
            log("[灵敏度RB] %s/%s: UL %d@%d, DL满RB=%d → UDCH(User Def)"
                "（表 7.3.5-2 RBstart>0，RMC 的 RBP:UL 仅 LOW|HIGH 无法表达；"
                "需 CMW-KS510）源=%s" %
                (band_label(band), bw_label(bw), lcrb, start,
                 cmds["dl_rb"], cmds["source"]))
        w("CONF:%s:CONN:UET RMC" % p)
        w(cmds["set_udch"])
        w(cmds["dl"])
        w(cmds["ul"])
    else:
        # RBstart=0：RMC 即可精确表达，与 ASRrun 参考流程完全一致
        mode = "RMC"
        cmds["rb_applied"] = (lcrb, 0)
        if log:
            log("[灵敏度RB] %s/%s: UL N%d@0, DL满RB=%d → RMC"
                "（表 7.3.5-2 RBstart=0，RMC:RBP:UL LOW 即可；TX full RB 独立）"
                "源=%s" %
                (band_label(band), bw_label(bw), lcrb, cmds["dl_rb"], cmds["source"]))
        w("CONF:%s:CONN:UET RMC" % p)
        w("CONF:%s:CONN:STYP RMC" % p)
        w("CONF:%s:CONN:RMC:DL N%d,%s,KEEP" % (p, cmds["dl_rb"], cfg["dl_mod"]))
        w("CONF:%s:CONN:RMC:UL N%d,QPSK,KEEP" % (p, lcrb))
        w("CONF:%s:CONN:RMC:RBP:DL LOW" % p)
        w("CONF:%s:CONN:RMC:RBP:UL LOW" % p)
    cmds["mode"] = mode
    if recheck:
        try:
            recheck(scpi, log, cfg=cfg)
        except TypeError:
            recheck(scpi, log)  # 兼容旧的两参数回调
    return cmds


# Full Cell BW Power 换算（CMW500 手册 23.3.6, PDF p.1708）：
#   "Full Cell BW Power resulting from the RS EPRE is displayed; calculated assuming
#    that the full DL cell bandwidth is used (all subcarriers) and all power offsets equal 0 dB."
#   CMW 无直接查询该值的 SCPI 命令，故软件换算：P_full = P_RS + 10log10(12*N_RB)。
#   N_RB(满) = 6/15/25/50/75/100 对应 1.4/3/5/10/15/20MHz。
BW_FULL_RB = {"B014": 6, "B030": 15, "B050": 25,
              "B100": 50, "B150": 75, "B200": 100}


def full_cell_bw_offset_db(bw):
    """RS EPRE(dBm/15kHz) -> Full Cell BW Pow(dBm) 的换算偏置 dB。"""
    rb = BW_FULL_RB.get(bw)
    if rb is None:
        mhz = BW_TO_MHZ.get(bw, 10.0)
        rb = max(1, int(round(mhz * 5)))
        if mhz <= 1.4:
            rb = 6
    return 10.0 * math.log10(12.0 * rb)


def epre_to_fullbw(level, bw):
    """把 RS EPRE 电平换算为 Full Cell BW Pow（dBm）；level 为 None 时原样返回。"""
    if level is None:
        return None
    try:
        return float(level) + full_cell_bw_offset_db(bw)
    except (TypeError, ValueError):
        return None


_REFSENS_CACHE = None


def refsens_table():
    """加载 3GPP REFSENS {band: {bwMHz: FullBW_dBm}}，本地 specs/LTE/refsens.json。

    键归一同 ul_cfg_table：频段 OB1/B1 -> "1"，带宽 10MHz -> 10.0。
    缺失/损坏返回空字典。
    """
    global _REFSENS_CACHE
    if _REFSENS_CACHE is not None:
        return _REFSENS_CACHE
    out = {}
    try:
        with open(LOCAL_REFSENS_JSON, encoding="utf-8") as f:
            raw = json.load(f)
        for band, bws in raw.items():
            row = {}
            for bw, v in (bws or {}).items():
                try:
                    row[float(bw)] = float(v)
                except (TypeError, ValueError):
                    continue
            if row:
                out[_band_num(band)] = row
    except Exception:
        out = {}
    _REFSENS_CACHE = out
    return out


def refsens_epre_for(band, bw):
    """(band,bw) 的理论 REFSENS 折算成 RS EPRE(dBm)；无数据返回 None。

    refsens.json 是 Full Cell BW Pow 口径（dBm），而仪器扫描用 RS EPRE，
    故减去 full_cell_bw_offset_db(bw)。真机验证误差 <1dB，可作首信道种子。
    """
    mhz = BW_TO_MHZ.get(bw)
    if mhz is None:
        return None
    fullbw = (refsens_table().get(_band_num(band)) or {}).get(mhz)
    if fullbw is None:
        return None
    return fullbw - full_cell_bw_offset_db(bw)


def refsens_limit_for(band, bw):
    """(band,bw) 的 3GPP 参考灵敏度限值 REFSENS（Full Cell BW Pow, dBm）；无数据返回 None。

    refsens.json 即 FullBW 口径，直接返回，无需换算（区别于 refsens_epre_for
    折算成 RS EPRE 作扫描种子）。bw 可为 'B100' 或 '10'/10.0。
    """
    mhz = BW_TO_MHZ.get(bw)
    if mhz is None:
        try:
            mhz = float(bw)
        except (TypeError, ValueError):
            return None
    return (refsens_table().get(_band_num(band)) or {}).get(mhz)


def refsens_limit_display(band, bw):
    """结果表显示串：3GPP 灵敏度 limit（FullBW dBm，保留 1 位）；缺失显示 '—'。"""
    v = refsens_limit_for(band, bw)
    return ("%.1f" % v) if v is not None else "—"


def refsens_limit_from_combo(label):
    """从组合显示名（如 'B1 / 10MHz / 低(CH50)'）解析 3GPP 灵敏度 limit 显示串。

    与 ul_rb_display_from_combo 同构：make_wide_row 只有 combo_label 时用。
    """
    parts = [p.strip() for p in str(label or "").split("/")]
    if len(parts) < 2:
        return "—"
    m = re.search(r"([\d.]+)\s*MHz", parts[1], re.I)
    if not m:
        return "—"
    return refsens_limit_display(parts[0], float(m.group(1)))


def build_bler_commands(cfg, level):
    """灵敏度探底：设置下行 RS EPRE 电平并触发一次 BLER 测量（手册 23.4.2 / 23.5.10）。

    判定量 = FETC:LTE:SIGN:EBLer:REL? 的第 2 值 Total Throughput(%)（idx 1）；
    Throughput ≥ sens_throughput（默认 95%）判 PASS，与 ASRrun 参考工具及
    TS 36.521-1 RefSens 判据一致（2026-11-25 实测核对修正，旧 idx4 已废弃）。

    REL 返回格式: Rel(0=OK), Total TPut(%), DTX(%), BLER(%), 副统计, 副统计
    （idx3 = 100 - idx1 恒成立；BLER 缺失 = 无统计块）。"""
    p = "LTE:SIGN"
    return {
        "set_level": "CONF:%s:DL:RSEPre:LEVel %.1f" % (p, level),
        "subframes": "CONF:%s:EBLer:SFRames %d" % (p, int(cfg["bler_subframes"])),
        "timeout": "CONF:%s:EBLer:TOUT %.1f" % (p, float(cfg["bler_timeout"])),
        "init": "INIT:%s:EBLer" % p,
        "state_q": "FETC:%s:EBLer:STAT?" % p,
        # ASRrun V2.0.7 参考工具用 REL 格式：idx 1 为 Total Throughput(%)（判定量）
        # 返回格式: Rel(0=OK), TotalTPut(%), DTX(%), BLER(%), 副统计, 副统计
        "fetch": "FETC:%s:EBLer:REL?" % p,
    }


def check_scpi_errors(scpi, log=None, ctx=""):
    """查询 SYST:ERR? 清空仪表错误队列并记录（对齐 ASRrun 参考工具做法）。

    每次配置写入后调用，及早发现 -203 (Command protected; option missing)
    之类的静默配置失败——参考工具 2026-11-25 log 正是靠它抓到 RMC:RBP:UL
    P25/P50 被 CMW 拒绝（缺 KS510 选件）。正常响应 "0,\"No error\""。
    返回错误列表（空=无错误）。"""
    errs = []
    for _ in range(8):
        try:
            resp = (scpi.query("SYST:ERR?", timeout=2.0) or "").strip()
        except Exception as e:
            if log:
                log("[SCPI错误] %s SYST:ERR? 查询失败：%s" % (ctx, type(e).__name__))
            return errs
        if not resp or resp.startswith("0,") or resp.upper().startswith('0,"NO'):
            break
        errs.append(resp)
    if errs and log:
        log("[SCPI错误] %s 检测到 %d 条仪表错误：%s" %
            (ctx, len(errs), "；".join(errs)))
    return errs


def build_psu_voltage_commands(cfg):
    """电源输出电压并开启（手册 E36311-90008 p.5 / p.69）。"""
    if not cfg.get("psu_enable") or not cfg.get("psu_voltage"):
        return []
    ch = cfg.get("psu_channel", "P6V")
    n = cfg.get("psu_channel_index", {ch: i + 1 for i, ch in enumerate(PSU_CHANNELS)}).get(ch, 1)
    v = float(cfg["psu_voltage"])
    return ["SOUR:VOLT %.4f,(@%d)" % (v, n), "OUTP ON,(@%d)" % n]


def read_psu_snapshot(psu, cfg, log=None):
    if psu is None or not psu.connected or not cfg.get("psu_enable"):
        return None
    ch = cfg.get("psu_channel") or PSU_CHANNELS[0]
    n = max(1, int(cfg.get("psu_samples", PSU_DEFAULT_SAMPLES)))
    currents, volts = [], []
    for i in range(n):
        try:
            resp = psu.query("MEAS:CURR? %s" % ch)
        except Exception as e:
            if log:
                log("[PSU] ⚠️ MEAS:CURR? 超时：%s" % type(e).__name__)
            return None   # PSU 通信异常，本次快照放弃
        if log:
            log("[PSU] MEAS:CURR? %s -> %s" % (ch, resp))
        v = parse_scalar(resp)
        if v is not None:
            currents.append(v)
        if cfg.get("psu_read_voltage", True):
            try:
                r2 = psu.query("MEAS:VOLT? %s" % ch)
            except Exception as e:
                if log:
                    log("[PSU] ⚠️ MEAS:VOLT? 超时：%s" % type(e).__name__)
                continue
            if log:
                log("[PSU] MEAS:VOLT? %s -> %s" % (ch, r2))
            v2 = parse_scalar(r2)
            if v2 is not None:
                volts.append(v2)
        if i < n - 1:
            time.sleep(0.03)
    if not currents:
        return None
    # E36312A MEAS:CURR? 默认返回 A，×1000 统一存为 mA（下游全用 mA 显示，避免处处换算）
    out = {
        "channel": ch, "current": 1000.0 * (sum(currents) / len(currents)),
        "current_min": 1000.0 * min(currents),
        "current_max": 1000.0 * max(currents),
        "samples": len(currents),
    }
    if volts:
        out["voltage"] = sum(volts) / len(volts)
        out["power"] = out["voltage"] * out["current"] / 1000.0  # V × mA / 1000 = W
    return out


def wait_for(scpi, query_cmd, predicate, timeout, interval=0.3, log=None,
             query_timeout=None, expect=None):
    """轮询状态查询直到 predicate 为真或超时；返回 (ok, last_resp)。

    query_timeout: 单次查询 socket 超时。仪表切频段重建小区期间会忙数秒，
    若沿用默认 10s，一次卡住的查询就吃掉整个轮询预算；短超时可在忙期继续重试。
    expect: 响应正则，透传 scpi.query 做语义校验（防流错位，见 RE_PSW 等）。
    """
    deadline = time.time() + max(0.1, timeout)
    last = ""
    while time.time() < deadline:
        try:
            last = scpi.query(query_cmd, timeout=query_timeout, expect=expect)
        except Exception as exc:
            if log:
                log("[轮询错误] %s" % exc)
            last = ""
        if log:
            log("[轮询] %s -> %s" % (query_cmd, last))
        if predicate(last):
            return True, last
        time.sleep(interval)
    return False, last


# 手册 PDF p.1738~1739: FETC:LTE:SIGN:PSW:STAT? 返回值
#   OFF(Signal Off) | ON(Signal On) | ATTached | CESTablished |
#   DISConnect | CONNecting | SIGNaling
# 关键: ON 只表示"小区信号已开"（generator started），**不代表 UE 注册**。
# 真正的 UE 注册/连接态是 ATTached / CESTablished。
# 注意：仪表固件实际返回 3 字母缩写（ATT / SIGN / CEST），而非手册全称，
# 因此全称与缩写都要认，否则会把已注册的 ATT 误判为未注册。
CONN_READY_STATES = {"ATT", "ATTACHED", "CEST", "CESTABLISHED"}


def conn_ready(state_text):
    """PSW:STAT? 是否处于真正的 UE 注册/连接态（而非仅 Signal On）。

    手册全称 ATTached/CESTablished 与仪表缩写 ATT/CEST 均视为已注册；
    ON/SIGN/OFF/CONNecting/DISConnect 不算。按 token 精确匹配，避免子串误判。
    """
    tokens = [t for t in re.split(r"[^A-Z]+", (state_text or "").upper()) if t]
    return any(t in CONN_READY_STATES for t in tokens)


def cell_ready(state_text):
    """CELL:STAT:ALL? 是否 ON,ADJ（信号真正可用；ON,PEND 尚不可用）。"""
    s = (state_text or "").upper()
    return "ON" in s and ("ADJ" in s or "PEND" not in s)


# 测量阶段复查注册态用的状态查询（与初始化 conn_state_q 同源）
CONN_RECHECK_QUERY = "FETC:LTE:SIGN:PSW:STAT?"

# --- 响应语义校验（防 SCPI 流错位，见 ScpiTransport.query 的 expect 参数）---
# PSW:STAT? 只返回单 token 状态，必须锚定整行：错位事故里 PSW 读到 CELL 的
# "ON,ADJ"，前缀 ^ON 同样能匹配，必须 $ 锚定才能识别错位。
RE_PSW = r"^(OFF|ON|ATT|CEST|CONN|DISC|SIGN)\s*$"
# CELL:STAT:ALL? 必然是 ON,xxx（PEND/ADJ）或 OFF
RE_CELL = r"^(ON,|OFF)"
# EBLer/MEV 状态查询：RUN/RDY 单 token
RE_STAT = r"^(RUN|RDY)\s*$"
# EBLer:REL? / MEV 数值查询：首字符必为数字或符号
RE_VALUE = r"^[-+0-9]"

# AT 自动恢复（掉线重注册加速）：与 GUI 手动 AT 发送共用同一把串口锁
_AT_SER_LOCK = threading.Lock()
AT_PORT_BAUD = 115200


def _at_recovery_port(cfg):
    """自动恢复可用的 AT 口：显式配置优先，其次 GUI 启动时探测到的口；无则 None。

    cfg["at_auto_recover"] = "0" 可整体禁用（默认 "1" 启用）。
    """
    if not serial:
        return None
    try:
        flag = str((cfg or {}).get("at_auto_recover", "1") or "1").strip().lower()
    except Exception:
        flag = "1"
    if flag in ("0", "false", "off"):
        return None
    port = str((cfg or {}).get("at_port") or "").strip()
    if not port or port.lower() == "auto":
        port = str((cfg or {}).get("at_port_detected") or "").strip()
    return port or None


def at_cfun_recover(port, log=None, settle=2.0):
    """下发 AT+CFUN=0 → settle → AT+CFUN=1，强制模组重启协议栈重搜网。

    强信号小区已在网侧就绪，CFUN=1 后 UE 全网扫频会立即注册到仪表小区，
    比被动等自主重选快且确定。单条失败不中断（CFUN=0 失败时 CFUN=1 仍可能
    救活）。返回 True=CFUN=1 已送达（收到 OK），False=未生效。
    """
    ser = None
    try:
        for attempt in range(2):
            try:
                ser = serial.Serial(port, AT_PORT_BAUD, timeout=1.0)
                break
            except Exception as e:
                if log:
                    log("[AT恢复] %s 打不开: %s" % (port, e))
                if attempt == 1:
                    return False
                time.sleep(0.8)
        ser.dtr = True
        ser.rts = True   # 确保唤醒（部分模组 DTR 未置位不响应 AT）
        time.sleep(0.3)

        def one(cmd, wait=8.0):
            ser.reset_input_buffer()
            ser.write((cmd + "\r\n").encode())
            t0 = time.time()
            got = ""
            while time.time() - t0 < wait:
                data = ser.readline()
                if not data:
                    continue
                line = data.decode(errors="replace").strip()
                if line:
                    if log:
                        log("[AT恢复] %s < %s" % (port, line))
                    got += line
                    if line in ("OK", "ERROR") or line.startswith("+CME ERROR"):
                        break
            return "OK" in got

        if log:
            log("[AT恢复] %s > AT+CFUN=0（关RF，UE 脱网）" % port)
        one("AT+CFUN=0")
        time.sleep(settle)
        if log:
            log("[AT恢复] %s > AT+CFUN=1（重启协议栈，重搜网注册仪表小区）" % port)
        return one("AT+CFUN=1", wait=10.0)
    except Exception as e:
        if log:
            log("[AT恢复] 异常: %s %s" % (type(e).__name__, e))
        return False
    finally:
        if ser is not None:
            try:
                ser.close()
            except Exception:
                pass


def at_recover_ue(scpi, cfg, log=None):
    """掉线后用 AT CFUN 强制 UE 重注册；成功返回注册态字符串，失败返回 None。

    流程：CFUN=0→1 → 等待回网（cfg["at_recover_wait"]，默认 60s；SIGN→ON
    抖动 3 轮即提前判负）→
    若只到 ATT（注册未连接）补 ACT CONN 进 CEST（TX 测量需要连接态）。
    失败时调用方回退原有重建小区路径，行为不劣化。
    """
    port = _at_recovery_port(cfg)
    if not port:
        return None
    if log:
        log("[AT恢复] 用 AT 口 %s 强制 UE 重注册（CFUN=0→1，替代被动等待）" % port)
    ok = False
    if _AT_SER_LOCK.acquire(timeout=5):
        try:
            ok = at_cfun_recover(port, log)
        finally:
            _AT_SER_LOCK.release()
    if not ok:
        if log:
            log("[AT恢复] CFUN 序列未送达，放弃 AT 恢复（回退重建小区）")
        return None
    try:
        _w = float((cfg or {}).get("at_recover_wait", 60) or 60)
    except (TypeError, ValueError):
        _w = 60.0
    _w = max(5.0, min(120.0, _w))
    # 自定义轮询（替代 wait_for）：除超时外，检测 UE 反复附着失败（SIGN→ON
    # 抖动 ≥3 轮）提前判负回退重建——真机 180539.log 组合 3：CFUN=1 后 UE 在
    # ON/SIGN 间抖动 30s 始终没到 ATT，重建小区反而 9s 恢复，死等只是浪费。
    deadline = time.time() + _w
    st = ""
    _flaps = 0
    _seen_sign = False
    while time.time() < deadline:
        try:
            st = scpi.query(CONN_RECHECK_QUERY, timeout=2.0, expect=RE_PSW) or ""
        except Exception:
            st = ""
        if log:
            log("[轮询] %s -> %s" % (CONN_RECHECK_QUERY, st))
        if conn_ready(st):
            break
        _su = (st or "").strip().upper()
        if _su == "SIGN":
            _seen_sign = True
        elif _su == "ON" and _seen_sign:
            _flaps += 1
            _seen_sign = False
            if _flaps >= 3:
                if log:
                    log("[AT恢复] UE 已 %d 次附着尝试失败（SIGN→ON 抖动），"
                        "提前放弃等待 → 回退重建小区" % _flaps)
                break
        time.sleep(0.5)
    if not conn_ready(st):
        if log:
            log("[AT恢复] CFUN=1 后未回网（PSW=%s，%.0fs 窗口）→ 回退重建小区"
                % ((st or "").strip() or "无响应", _w))
        return None
    if "CEST" not in (st or "").upper():
        scpi.write("CALL:LTE:SIGN:PSW:ACT CONN")
        if log:
            log("[SCPI] CALL:LTE:SIGN:PSW:ACT CONN（ATT→CEST 建立连接）")
        ok3, st = wait_for(scpi, CONN_RECHECK_QUERY, conn_ready,
                           8.0, 0.5, log, query_timeout=2.0, expect=RE_PSW)
        if not ok3:
            st = "ATT"   # 已注册即可继续，测前复查仍有保护
    if log:
        log("[AT恢复] UE 已重注册（%s）" % st)
    return st


def ensure_registered(scpi, log=None, timeout=8.0, interval=0.5):
    """每个组合测量前复查 UE 是否仍处于注册/连接态。

    关键：测量循环里每个组合都会重新配置 BAND/带宽/CHAN:DL，这些信令小区
    参数的切换可能使已注册的 UE 掉线。若不复查就直接测，会得到一堆 INV
    无效值还跑完全部组合（"没注册就测试了"）。
      - 已注册：立即返回，零额外开销；
      - 未注册：在 timeout 内等待 UE 重新注册；
      - 超时仍未注册：抛 ScpiError 中止整个测量，不产出无效数据/空报告。
    """
    st = ""
    try:
        st = scpi.query(CONN_RECHECK_QUERY, expect=RE_PSW)
    except Exception as exc:
        if log:
            log("[注册复查] 状态查询异常: %s" % exc)
        st = ""
    if conn_ready(st):
        if log:
            log("[注册复查] %s：UE 已注册，允许测量" % st)
        return st
    if log:
        log("[注册复查] %s：UE 未注册，等待重新注册（最多 %.0fs）..." % (st, timeout))
    ok, st = wait_for(scpi, CONN_RECHECK_QUERY, conn_ready, timeout, interval, log,
                      expect=RE_PSW)
    if not ok:
        raise ScpiError(
            "组合测量前 UE 未注册到小区（PSW:STAT=%s）。\n"
            "切换频段/带宽/信道可能使 UE 掉线，测量已中止以避免产出无效数据。\n"
            "请确认：① DUT 在线且信号足够；② 切换参数后 UE 能自动重新注册；"
            "③ 必要时每个频点单独重新初始化。" % st)
    if log:
        log("[注册复查] %s：UE 已重新注册" % st)
    return st


def reconnect_after_styp(scpi, log=None, timeout=20.0, cfg=None, recovery_level=None,
                         retune_ctx=None):
    """切 STYP（RMC↔UDCH）后主动触发 UE 重新附着。

    真机 log（cmw500_20261123_093433）证明：切信道后 UE 仍 ATT，但随后切
    STYP UDCH 会从 ATT 掉到 ON，且纯被动等待不会自动重注册（与换频段掉线
    同理）。按手册连接建立顺序，掉线时主动 CALL:PSW:ACT CONN 重新触发附着，
    短查询超时轮询等 ATT/CEST；已注册则零开销返回。

    掉线重附着前必须先把 DL RS EPRE 拉回强信号：灵敏度扫描后电平停在
    -110~-125 dBm，UE 在弱信号下测不到小区、无法完成随机接入，此时直接发
    CALL:PSW:ACT CONN 会一直停在 ON 直到超时（124004.log 已证实）。恢复电平
    取 cfg["rsrp"]（初始化注册电平，默认 -85），等 UE 感知强信号后再附着。
    已注册时零开销返回、不碰电平——避免破坏正在进行的灵敏度扫描。

    快速升级策略（方案 A）：真机 144321.log 显示软掉线的 ACT CONN 在 1 秒内
    就进 CEST，硬掉线则 20 秒全程 PSW:STAT=ON。因此 ACT CONN 后先用
    cfg["attach_quick_timeout"]（默认 5s）短窗口判定，未进 ATT/CEST 且提供
    了 retune_ctx（{"band":..,"bw":..,"earfcn":..}）时直接调用 retune_cell
    重建小区，不再死等 timeout×2（该场景原耗时约 50s）。未提供 retune_ctx
    的调用点（apply_sens_rb 的 recheck）继续用满 timeout，总超时不变。

    cfg / recovery_level 为可选参数：apply_sens_rb 的 recheck 回调只传
    (scpi, log)，此时按默认 -85 恢复；显式传 cfg 时用其 rsrp。
    """
    p = "LTE:SIGN"

    def w(cmd):
        scpi.write(cmd)
        if log:
            log("[SCPI] " + cmd)

    # 首次查询：容忍一次超时（仪表在切换弱信号后可能忙数秒不响应）。真机
    # 152654.log 组合 3 曾因单次查询超时得到空状态、被当掉线而白耗 18s 重建。
    # 用 wait_for 短窗口轮询，恢复即零开销返回，不会误判。
    _pre_ok, st = wait_for(scpi, CONN_RECHECK_QUERY, conn_ready,
                           2.0, 0.5, log, query_timeout=2.0, expect=RE_PSW)
    if _pre_ok:
        if log:
            log("[STYP重连] %s：UE 已注册，无需重附着" % st)
        return st

    # 掉线：先拉回强信号（注册电平），让 UE 能检测小区并完成随机接入
    if recovery_level is None:
        try:
            recovery_level = float((cfg or {}).get("rsrp", -85) or -85)
        except (TypeError, ValueError):
            recovery_level = -85.0
    w("CONF:%s:DL:RSEPre:LEVel %.1f" % (p, recovery_level))
    if log:
        log("[重附着] 先把 RS EPRE 拉回强信号 %.1f dBm（注册电平），"
            "避免灵敏度弱信号下 UE 无法随机接入" % recovery_level)
    time.sleep(0.8)  # 让 UE 完成强信号小区检测/测量后再触发附着

    if log:
        log("[STYP重连] %s：UE 掉线，主动触发 CALL:PSW:ACT CONN 重新附着" % st)
    w("CALL:%s:PSW:ACT CONN" % p)
    try:
        _quick = float((cfg or {}).get("attach_quick_timeout", 5) or 5)
    except (TypeError, ValueError):
        _quick = 5.0
    _quick = max(1.0, min(_quick, timeout))
    ok, st = wait_for(scpi, CONN_RECHECK_QUERY, conn_ready,
                      _quick, 0.5, log, query_timeout=2.0, expect=RE_PSW)
    if not ok and retune_ctx:
        # 硬掉线：短窗口内 PSW 一直 ON，软等待无意义。
        # 先试 AT CFUN=0→1 强制重注册（强信号小区已在网侧就绪，UE 全网扫频
        # 立即回网，比重建小区快）；AT 不可用/未生效再回退重建小区。
        _st_at = at_recover_ue(scpi, cfg, log) if cfg else None
        if _st_at:
            return _st_at
        if log:
            log("[重附着] ACT CONN %.1fs 内 PSW:STAT=%s 未进 ATT/CEST，"
                "判定硬掉线 → 直接重建小区（不再死等 %.0fs×2）" % (_quick, st, timeout))
        try:
            retune_cell(scpi, cfg or {}, retune_ctx.get("band"),
                        retune_ctx.get("bw"), retune_ctx.get("earfcn"), log)
            try:
                st = scpi.query(CONN_RECHECK_QUERY, timeout=2.0, expect=RE_PSW)
            except Exception:
                st = "ATT"
            if log:
                log("[重附着] 重建小区后 UE 已恢复（%s）" % st)
            return st
        except Exception as exc:
            raise ScpiError("ACT CONN 与重建小区均未能恢复 UE 附着（%s）。" % exc)
    if not ok:
        # 未提供 retune_ctx（如切 STYP 软掉线）：补满剩余超时，总时长与旧版一致
        _remain = max(0.0, timeout - _quick)
        if _remain > 0:
            ok, st = wait_for(scpi, CONN_RECHECK_QUERY, conn_ready,
                              _remain, 0.5, log, query_timeout=2.0, expect=RE_PSW)
    if not ok:
        raise ScpiError("切 STYP 后 UE 未重新注册（PSW:STAT=%s）。" % st)
    if log:
        log("[STYP重连] UE 已重新附着（%s）" % st)
    return st


# ----------------------------------------------------------------------------
# 换频段/带宽：快速直切（优先） + 失败回退重建小区
# ----------------------------------------------------------------------------


def _apply_cell_params(scpi, cfg, band, bw, earfcn, w, log=None, skip_band=False,
                       order=None, park_earfcn=None):
    """下发频段/带宽/信道 + 按新带宽重配 RMC RB + 动态 EATT（快切/重建共用）。

    关键（真机 2026-11-23 19:20）：带宽切换后 RMC 必须按新带宽重配 DL/UL RB，
    否则小区是 15MHz(B150) 而 RMC 还是 N50(10MHz)，UE attach 时 DL SCH RB 与
    小区带宽不匹配 → 反复 SIGN 后放弃，永远不到 ATT。

    skip_band: 同频段热切带宽时跳过 BAND 写（写不写不影响仪表，跳过可减少
    连接态下的多余 CONF 写入；参考工具同频段热切只写 BANDwidth/RMC）。

    order: 写入顺序，默认 ("band","bw","chan","rmc","eatt")。热切带宽时由
    try_fast_retune 按带宽增大/减小方向重排（见其 docstring）——顺序错了
    会产生非法中间态（新带宽 + 旧信道越界 / 新带宽 + 旧 RMC RB 超限），
    仪表拒绝写入导致小区坏掉、UE 掉线（真机 191624.log 三次带宽切换实测）。

    park_earfcn: "park" 步骤写入的停靠信道（带宽增大方向先用中间信道停靠，
    见 try_fast_retune；None 或步骤未用到则不写）。
    """
    p = "LTE:SIGN"

    def _band():
        if band and not skip_band:
            w("CONF:%s:BAND %s" % (p, band))

    def _bw():
        if bw:
            w("CONF:%s:CELL:BAND:DL %s" % (p, bw))

    def _chan():
        if earfcn is not None and int(earfcn) >= 0:
            w(build_channel_cmd(cfg, earfcn))

    def _park():
        if park_earfcn is not None and int(park_earfcn) >= 0:
            w(build_channel_cmd(cfg, park_earfcn))

    def _rmc():
        if bw:
            _rb = int(BW_TO_MHZ.get(bw, 10) * 5)
            if bw == "B014":
                _rb = 6
            _dl_mod = (cfg.get("dl_mod") or "QPSK")
            w("CONF:%s:CONN:RMC:DL N%d,%s,KEEP" % (p, _rb, _dl_mod))
            w("CONF:%s:CONN:RMC:UL N%d,QPSK,KEEP" % (p, _rb))
            w("CONF:%s:CONN:RMC:RBP:DL LOW" % p)
            w("CONF:%s:CONN:RMC:RBP:UL LOW" % p)

    def _eatt():
        _mhz = earfcn_to_dl_center_mhz(band, earfcn)
        _loss = lookup_loss_mhz(cfg, _mhz) if _mhz else None
        if _loss is not None and _loss > 0:
            w("CONF:%s:RFS:EATT:OUTP %.1f" % (p, _loss))
            w("CONF:%s:RFS:EATT:INP %.1f" % (p, _loss))
            if log:
                log("[EATT] DL=%.1f MHz → 外部线缆损耗 %.1f dB（OUTP/INP 均补偿）"
                    % (_mhz, _loss))
        elif _loss == 0.0:
            w("CONF:%s:RFS:EATT:OUTP 0.0" % p)
            w("CONF:%s:RFS:EATT:INP 0.0" % p)

    _steps = {"band": _band, "bw": _bw, "chan": _chan, "park": _park,
              "rmc": _rmc, "eatt": _eatt}
    for _name in (order or ("band", "bw", "chan", "rmc", "eatt")):
        _steps[_name]()


def _band_mid_earfcn(band, bw):
    """某频段指定带宽的中间信道 EARFCN（信道表缺失/频段未知返回 None）。

    中信道由频段范围居中内缩算出，跨带宽恒为频段中心（B5=2525、B2=900），
    故旧带宽的中信道在新带宽下同样合法（区间嵌套），可作带宽切换的停靠点。
    """
    try:
        for _pos, _e in band_channels(band, bw):
            if _pos == "中":
                return int(_e)
    except Exception:
        pass
    return None


def try_fast_retune(scpi, cfg, band, bw, earfcn, log=None, last_cell=None):
    """频段/带宽切换的快速路径：强信号下连接态热改参数（不重启小区）。

    参考工具 log（2026-11-24-10-14-00_18340）决定性证据：连接态（CEST）下
    热写 CELL:BANDwidth:DL + RMC:UL（B100↔B030↔B050 双向 7 次），UE 全程
    保持 CEST 不掉线，热切后测量直接出有效值——参考工具换带宽后从不重发
    CELL ON（log 里唯一一次 CELL ON 是开小区的第一次 OFF→ON）。对已 ON
    小区重发 CELL ON 会触发小区整机重启：仪表 RF 全重调（查询黑洞 ~20s）
    + UE 强制掉线 + 重新注册，整条恢复链 50~60s（真机 191624.log 三次带宽
    切换实测，其中 ~21s 忙期 + 8s 白等 + ~20s AT 恢复）。

    流程：热写参数（带宽 + 按新带宽重配 RMC + EATT；换频段才写 BAND 并保留
    CELL ON——不同频率 RF 本就要重调、UE 本就要重选，2026-11-23 有掉线实证）
    → 短窗口（cfg["fast_retune_timeout"]，默认 8s）轮询 PSW:STAT；进
    ATT/CEST 即返回 True（UE 保持注册，热切成功，~3-5s）。

    带宽增大方向的两段式（2026-11-25 用户方案，B5/B2 两轮 12+18 组合实测）：
    联合改「信道+带宽」在 →10MHz 必触发仪表 ~10-12s 查询黑洞、UE 必掉线
    （1.4→3/3→5/10→15/15→20 均忙期 3-5s 热切成功）。改为：
      1. 停靠中间信道 CH_mid（同带宽信道切换，实测干净；中信道跨带宽恒为
         频段中心，B5=2525、B2=900，旧/新带宽下均合法）；
      2. 确认窗口等 UE 回 ATT/CEST（fast_retune_timeout）；
      3. 同中心纯带宽切换 + 新 RMC + 目标信道（参考工具实证不掉线的正是
         同中心改带宽模式）。
    首窗口未确认时（可能仍在忙期）：cfg["fast_retune_busy_wait"]（默认 25s）
    宽限确认——合法 ATT/CEST 响应视为 UE 自恢复成功；合法 ON/OFF/DISC 才判
    真掉线（避免 CFUN=0 打断自主重附着）。掉线后分流：
      - AT 口可用 → 返回 False 交调用方 AT 恢复（CFUN=0→1 ~20s，确定）；
      - 无 AT 口 → 被动等 UE 自主重附着（cfg["fast_retune_passive_wait"]，
        默认 30s；参考工具实测自主重附着 ~26s）→ 仍未回 False 由调用方
        retune_cell 重建兜底。

    写入顺序（真机 191624.log + 信道表证实的关键）：带宽切换必然同时换中心
    信道（每带宽的低/中/高信道不同），若按默认顺序先带宽后信道，中间态
    「新带宽 + 旧信道」可能越界——B5 实例：1.4M 低信道 869.7 MHz < 3M 合法
    下限 870.4（越界 0.7 MHz）；5M→10M 时 871.5 < 874.0（越界 2.5 MHz）。
    仪表拒绝该写入 → 小区坏掉 → UE 掉线。故按方向重排（合法区间随带宽变宽
    而变宽，嵌套关系保证下述中间态恒合法）：
      - 带宽增大：先停靠中间信道（见上），兜底路径为 信道 → 带宽 → RMC；
      - 带宽减小：RMC → 带宽 → 信道（先降 RB 避免旧 N 超新带宽上限；旧信道
        必在变宽后的合法范围内）。
    """
    p = "LTE:SIGN"

    def w(cmd):
        scpi.write(cmd)
        if log:
            log("[SCPI] " + cmd)

    if log:
        log("[频段快切] %s/%s CH%s：连接态热改参数（不重启小区，参考工具模式）"
            % (band_label(band), bw_label(bw), earfcn))

    band_changed = (not last_cell) or (not last_cell[0]) or (last_cell[0] != band)
    old_bw = (last_cell[1] if last_cell and len(last_cell) > 1 else "") or ""
    _mhz_new = BW_TO_MHZ.get(bw or "")
    _mhz_old = BW_TO_MHZ.get(old_bw)
    _increase = ((not band_changed) and bool(old_bw)
                 and _mhz_new is not None and _mhz_old is not None
                 and _mhz_new > _mhz_old)
    try:
        _target = int(earfcn)
    except (TypeError, ValueError):
        _target = -1

    try:
        _fast = float((cfg or {}).get("fast_retune_timeout", 5) or 5)
    except (TypeError, ValueError):
        _fast = 5.0
    _fast = max(1.0, min(_fast, 20.0))

    # 带宽增大方向（两段式，2026-11-25 用户方案）：先停靠中间信道（同带宽
    # 信道切换，实测干净 ~3-4s），等 UE 确认回网后再做「同中心纯带宽切换」
    # ——参考工具实证不掉线的正是这种同中心改带宽模式。两轮实测（B5/B2）
    # 5→10MHz 联合改「信道+带宽」必触发仪表 ~10-12s 忙期、UE 必掉线，而
    # 1.4→3/3→5/10→15/15→20 均忙期 3-5s 热切成功；拆成两步、中间留确认
    # 窗口，让仪表把两次变更分开处理。中间信道跨带宽恒同（B5=2525、B2=900），
    # 停靠点在旧带宽合法、新带宽下亦合法（区间嵌套）。
    _parked = False
    if _increase:
        _mid = _band_mid_earfcn(band, old_bw)
        if _mid is not None and _mid != _target:
            if log:
                log("[频段快切] 带宽增大：先停靠中间信道 CH%d（同带宽信道切换），"
                    "UE 回网后再做同中心带宽切换" % _mid)
            _apply_cell_params(scpi, cfg, band, bw, earfcn, w, log,
                               skip_band=True, order=("band", "park"),
                               park_earfcn=_mid)
            ok_park, st_park = wait_for(scpi, "FETC:%s:PSW:STAT?" % p,
                                        conn_ready, _fast, 0.5, log,
                                        query_timeout=2.0, expect=RE_PSW)
            if log:
                log("[频段快切] 停靠确认：%s"
                    % ("UE 已回（%s），开始切换带宽" % st_park
                       if ok_park else "UE 未确认（PSW=%s），继续切换带宽"
                       % (st_park or "忙")))
            _parked = True
    if _parked:
        # 第二段：同中心纯带宽切换 + 新带宽 RMC + 目标信道 + EATT
        _apply_cell_params(scpi, cfg, band, bw, earfcn, w, log,
                           skip_band=True,
                           order=("bw", "rmc", "chan", "eatt"))
    elif band_changed or not old_bw or _mhz_new is None or _mhz_old is None \
            or _mhz_new == _mhz_old:
        order = ("band", "bw", "chan", "rmc", "eatt")
    elif _mhz_new > _mhz_old:
        # 带宽增大（无中信道可用时的兜底）：先对齐目标信道（必在旧带宽范围内），
        # 再扩带宽，最后升 RMC
        order = ("band", "chan", "bw", "rmc", "eatt")
        if log:
            log("[频段快切] 带宽增大：写入顺序 信道→带宽→RMC"
                "（避免中间态信道超出新带宽合法范围）")
    else:
        # 带宽减小：先降 RMC（旧带宽内容合法），再收带宽（旧信道必在变宽的
        # 合法范围内），最后对齐信道
        order = ("band", "rmc", "bw", "chan", "eatt")
        if log:
            log("[频段快切] 带宽减小：写入顺序 RMC→带宽→信道"
                "（避免中间态 RMC RB 超出带宽上限）")
    if not _parked:
        _apply_cell_params(scpi, cfg, band, bw, earfcn, w, log,
                           skip_band=not band_changed, order=order)

    # 换频段才重发 CELL ON（触发 RF 重调）；同频段带宽热切绝不重启小区
    if band_changed:
        w("SOUR:%s:CELL:STAT ON" % p)

    ok, st = wait_for(scpi, "FETC:%s:PSW:STAT?" % p, conn_ready,
                      _fast, 0.5, log, query_timeout=2.0, expect=RE_PSW)
    if ok:
        if log:
            log("[频段快切] %s/%s CH%s 热切成功（UE %s 保持），未重启小区"
                % (band_label(band), bw_label(bw), earfcn, st))
        return True

    # 首窗口未确认（PSW=%s）：可能仍在仪表忙期内（→10MHz 实测查询黑洞
    # ~10-12s > 首窗口 8s），超时/空响应并不能证明 UE 掉线。等忙期结束后
    # 再给一次确认窗口：合法 ATT/CEST 响应 → UE 自恢复成功；合法 ON/OFF/DISC
    # → 真掉线才交 AT（避免 CFUN=0 打断 UE 正在进行的自主恢复——B5/B2 两轮
    # 实测 CFUN=0 响应慢 11s 正是"模组重附着中途被打断"）；SIGN/CONN 属
    # 注册进行中，继续等。
    try:
        _busy_wait = float((cfg or {}).get("fast_retune_busy_wait", 25) or 25)
    except (TypeError, ValueError):
        _busy_wait = 25.0
    _busy_wait = max(5.0, min(_busy_wait, 60.0))
    _busy_deadline = time.time() + _busy_wait
    _st2 = ""
    while time.time() < _busy_deadline:
        try:
            _resp = (scpi.query("FETC:%s:PSW:STAT?" % p, timeout=2.0)
                     or "").strip()
        except Exception:
            _resp = ""          # 仍在忙期（查询黑洞）
        if _resp and re.match(RE_PSW, _resp):
            if conn_ready(_resp):
                if log:
                    log("[频段快切] %s/%s CH%s 热切成功（忙期结束 UE 自恢复 %s），"
                        "未重启小区" % (band_label(band), bw_label(bw), earfcn,
                                       _resp))
                return True
            if _resp in ("ON", "OFF", "DISC"):
                _st2 = _resp    # 真掉线的确定性证据
                break
            # SIGN/CONN：注册进行中，继续等到窗口耗尽
        time.sleep(0.5)
    st = _st2 or st

    # 热切后 UE 掉线（PSW=ON）：AT 可用立即分流给调用方 AT 恢复；
    # 无 AT 口则被动等自主重附着（参考工具实测 ~26s）。
    if _at_recovery_port(cfg):
        if log:
            log("[频段快切] 热切后 UE 掉线（PSW=%s）→ 交由 AT 恢复" % st)
        return False
    try:
        _passive = float((cfg or {}).get("fast_retune_passive_wait", 30) or 30)
    except (TypeError, ValueError):
        _passive = 30.0
    _passive = max(5.0, min(_passive, 60.0))
    if log:
        log("[频段快切] 热切后 UE 掉线（PSW=%s）且无 AT 口，被动等自主重附着"
            "（最多 %.0fs）" % (st, _passive))
    ok, st = wait_for(scpi, "FETC:%s:PSW:STAT?" % p, conn_ready,
                      _passive, 0.5, log, query_timeout=2.0, expect=RE_PSW)
    if ok:
        if log:
            log("[频段快切] 热切后 UE 自主重附着（%s），跳过重建小区" % st)
        return True
    if log:
        log("[频段快切] %.0fs 内 UE 未回（PSW=%s）→ 回退重建小区"
            % (_passive, st))
    return False


def retune_cell(scpi, cfg, band, bw, earfcn, log=None):
    """切换到新频段/带宽时重建小区并让 UE 重新附着。

    真机 log（2026-11-23）证实：连接态直接写 CONF:LTE:SIGN:BAND 换频段，
    仪表会忙约 10s 不响应 PSW:STAT? 查询，UE 从 ATT 掉到 ON 且不会自动
    重新注册，纯被动等待无效。必须按手册 p.1732/1733 的连接建立顺序重来：
      CELL OFF → 重配 BAND/带宽/信道 → CELL ON → 等 ON,ADJ
      → CALL:PSW:ACT CONN → 等 UE 重新 ATT/CEST。
    """
    p = "LTE:SIGN"

    def w(cmd):
        scpi.write(cmd)
        if log:
            log("[SCPI] " + cmd)

    if log:
        log("[频段切换] 重建小区：%s/%s CH%s"
            "（关小区→重配→开小区→等 ON,ADJ→UE 重新附着）" %
            (band_label(band), bw_label(bw), earfcn))

    # ① 关小区：让 UE 明确去附着，并释放旧频段硬件资源
    w("SOUR:%s:CELL:STAT OFF" % p)
    time.sleep(0.3)  # 等小区真正释放，再改频段（避免仪表忙期写配置）

    # ② 先切 duplex（TDD 需 DMODe TDD + ULDL/SSUB；切回 FDD 需 DMODe FDD）。
    # duplex 决定 BAND 允许范围，必须在写 BAND 之前下发（手册 p.306）。
    if is_tdd_band(band):
        for cmd in tdd_cell_commands(cfg, band):
            w(cmd)
    else:
        for cmd in fdd_cell_commands():
            w(cmd)

    # ②b 重配频段 / 带宽 / 信道 + 按新带宽重配 RMC RB + 动态 EATT
    # （与 try_fast_retune 共用，避免两处带宽 RB 换算逻辑漂移）
    _apply_cell_params(scpi, cfg, band, bw, earfcn, w, log)

    # ③ 开小区并等 ON,ADJ（短查询超时，忙期可继续轮询）
    w("SOUR:%s:CELL:STAT ON" % p)
    ok, st = wait_for(scpi, "SOUR:%s:CELL:STAT:ALL?" % p, cell_ready,
                      20.0, 0.5, log, query_timeout=2.0, expect=RE_CELL)
    if not ok:
        raise ScpiError("切换频段后小区未就绪（需 ON,ADJ）: %s" % st)

    # ③a 附着前把下行电平拉回强信号（注册电平 rsrp）：
    # 灵敏度测完重启小区的路径（measure_sensitivity 末尾 retune_cell）电平仍停在
    # -110~-125 dBm，弱信号下 UE 测不到小区、CALL:PSW:ACT CONN 会一直超时。
    # 切信道路径虽已在调用方做过"切换保护"，这里重复拉一次无副作用。
    try:
        _recover = float((cfg or {}).get("rsrp", -85) or -85)
    except (TypeError, ValueError):
        _recover = -85.0
    w("CONF:%s:DL:RSEPre:LEVel %.1f" % (p, _recover))
    if log:
        log("[频段切换] 附着前把 RS EPRE 拉回强信号 %.1f dBm（注册电平）" % _recover)
    time.sleep(0.5)  # UE 感知强信号

    # ④ 主动触发 UE 附着，等 ATT/CEST
    # 真机 BW 切换（10→15MHz）后 UE 偶发停在 ON 不进 ATT，
    # 两次 CALL:PSW:ACT CONN 重试 + 分段等待（首次重附着仪表忙期 10~20s）
    _last_st = None
    for _attempt in range(2):
        w("CALL:%s:PSW:ACT CONN" % p)
        ok, st = wait_for(scpi, "FETC:%s:PSW:STAT?" % p, conn_ready,
                          20.0, 0.5, log, query_timeout=2.0, expect=RE_PSW)
        _last_st = st
        if ok:
            break
        if log:
            log("[重附着] 第 %d 次 CALL:PSW:ACT CONN 后 PSW:STAT=%s，"
                "未进入 ATT/CEST，重试一次。" % (_attempt + 1, st))
    if not ok:
        # 容错：如果状态是 ON（小区已 ON,ADJ 只是 UE 没跟上），
        # 再等 15s 让 UE 被动附着；仍不行才真正放弃
        if _last_st and "ON" in _last_st.upper() and "ATT" not in _last_st.upper() \
                and "CEST" not in _last_st.upper():
            if log:
                log("[重附着] 小区 ON,ADJ 但 UE 停在 %s，再等 15s 观察被动附着…" % _last_st)
            ok, st = wait_for(scpi, "FETC:%s:PSW:STAT?" % p, conn_ready,
                              15.0, 0.5, log, query_timeout=2.0, expect=RE_PSW)
            _last_st = st
        if not ok:
            raise ScpiError("切换频段/带宽后 UE 未重新附着（PSW:STAT=%s）。" % _last_st)

    if log:
        log("[频段切换] %s/%s 小区 ON,ADJ，UE 已重新附着（%s）" %
            (band_label(band), bw_label(bw), st))
    # ⑤ 上行稳定等待：ATT 只表示 UE 已附着，尚未必被调度上行发射。真机
    #    134654.log 组合 7 切到 5MHz 后刚 ATT 就 MAXPower + MEV，结果全 INV、
    #    Reliability=3 Overdriven 中止。等待后复查注册，确保交给 TX 的是
    #    已稳定在网、正在上行的 UE（不新增 SCPI 查询开销以外的测量）。
    try:
        _asettle = max(0.0, float(cfg.get("attach_settle", 2) or 2))
    except (TypeError, ValueError):
        _asettle = 2.0
    if _asettle:
        if log:
            log("[频段切换] UE 已 ATT，等待 %.1fs 让上行调度稳定后再测 TX" % _asettle)
        time.sleep(_asettle)
        try:
            _st_final = scpi.query(CONN_RECHECK_QUERY, timeout=2.0, expect=RE_PSW)
        except Exception:
            _st_final = ""
        if not conn_ready(_st_final):
            raise ScpiError("切频段/带宽后等待 %.1fs 期间 UE 掉线（PSW:STAT=%s），"
                            "已中止以避免后续 TX 无效数据。" % (_asettle, _st_final))
        if log:
            log("[频段切换] 上行稳定等待后 UE 仍在网（%s）" % _st_final)
    return st


# ----------------------------------------------------------------------------
# 测量引擎（信令模式）
# ----------------------------------------------------------------------------


def measure_tx(scpi, cfg, enabled, log=None, band="", power_mode="high"):
    """连接态下测 UE 上行：功率 / ACLR / EVM / 频率误差（按 power_mode 档位）。

    关键（手册 PDF p.1510）: *RST 后 MEV 处于 OFF，必须先 INIT 触发一次
    single-shot，等状态进入 RDY 后再 FETCh，否则结果全是 NAV。

    手册 23.5.1.2 (p.1736): FETCh 返回数组首位是 Reliability Indicator，
    单独解析用于精确诊断（4=Underdriven 输入信号过低，本次实测日志即为此）。

    手册 23.2.3 (p.1681): 切换信道/频段/带宽后连接会进入 Signaling in
    Progress 重同步，UE 可能尚未被调度上行发射 → 首次结果全 INV。此时重新
    INIT 触发测量重试 mev_retry 次，给 UE 时间在新信道稳定并发射。
    """
    q = build_tx_queries(cfg)
    # 显式 MEV 配置：每次 TX 测量前重置，避免切信道/切 STYP 后状态残留
    # band 必须传入：TDD 的 MSUBframes 偏移为 2（FDD 为 0），见函数注释
    apply_mev_config(scpi, cfg, log, band=band)

    def query(cmd, timeout=None):
        # 按命令语义选 expect：状态查询（STAT?）回 RDY/RUN 单 token，
        # 其余为 MEV 数值/CSV——防止 SCPI 流错位后拿到错误响应当真值
        exp = RE_STAT if cmd.endswith(":STAT?") else RE_VALUE
        r = scpi.query(cmd, timeout=timeout, expect=exp)
        if log:
            log("%s -> %s" % (cmd, r))
        return r

    def write(cmd):
        scpi.write(cmd)
        if log:
            log("[SCPI] " + cmd)

    def _num(key, default, lo, hi, cast=float):
        try:
            return cast(cfg.get(key, default))
        except (TypeError, ValueError):
            return cast(default)

    need_mod = bool(enabled & {"maxpower", "evm", "freqerr"})
    need_aclr = "aclr" in enabled
    retry = max(0, min(5, int(_num("mev_retry", 2, 0, 5))))
    gap = max(0.0, min(10.0, _num("mev_retry_gap", 1.0, 0.0, 10.0)))
    # 保险：TX 测量前强制关 OCNG（灵敏度段异常退出可能残留 ON，令 UE 上行
    # 过驱、MEV 返回 Reliability=3 Overdriven）。幂等下发，无副作用。
    write("CONF:LTE:SIGN:DL:OCNG OFF")

    def one_shot():
        # 0) 先终止上一次测量，避免重触发时状态残留（ASRrun 同样每次 ABORt）
        write("ABORt:%s:MEV" % "LTE:MEAS")
        # 1) 触发一次单次测量（不 INIT 则 FETCh 只会返回 NAV）
        write(q["init"])
        # 2) 等待测量结束进入 RDY
        ok, state = wait_for(scpi, q["state"], lambda s: "RDY" in s.upper(),
                             20.0, 0.3, log, expect=RE_STAT)
        if not ok and log:
            log("[警告] MEV 20s 内未进入 RDY（当前=%s），仍尝试读取结果。" % state)
        # 3) 读取原始结果（保留原始串以解析首位 Reliability）
        raw_ma = query(q["mod_avg"]) if need_mod else ""
        raw_me = query(q["mod_ext"]) if need_mod else ""
        raw_ac = query(q["aclr"]) if need_aclr else ""
        rel_code, rel_text = parse_reliability(raw_ma or raw_ac or raw_me)
        # NCAP 占比：大片 NCAP（未捕获）= UE 根本没有上行发射，典型原因是
        # RMC/UDCH RB 配置与小区带宽失配（如小区按 10MHz 建立、RMC 仍 N50，
        # 组合却切到 1.4MHz）——真机 180539.log Run1 曾被误报 Overdriven。
        _ncap = 0
        _ntok = 0
        for _t in (raw_ma or "").split(",")[1:]:
            _t = _t.strip().upper()
            if _t:
                _ntok += 1
                if _t == "NCAP":
                    _ncap += 1
        d = {
            "state": state,
            "mev_ready": ok,
            "reliability": rel_code,
            "reliability_text": rel_text,
            "ncap_ratio": (_ncap / _ntok) if _ntok else 0.0,
            "mod_avg": parse_values(raw_ma) if need_mod else [],
            "mod_ext": parse_values(raw_me) if need_mod else [],
            "aclr": parse_values(raw_ac) if need_aclr else [],
            "queries": q,
        }
        d["raw"] = {k: ",".join("" if v is None else "%.6g" % v for v in d[k])
                    for k in ("mod_avg", "mod_ext", "aclr")}
        # 有效性判定：功率/EVM/频偏/ACLR 关键位若全为 None（INV/NAV）→ UE 未在发射
        checks = []
        if need_mod:
            checks += [at(d["mod_avg"], 17), at(d["mod_avg"], 2), at(d["mod_avg"], 15)]
        if need_aclr:
            checks += [at(d["aclr"], 3), at(d["aclr"], 5)]
        d["valid"] = any(v is not None for v in checks) if checks else True
        return d

    d = one_shot()
    attempt = 1
    while not d["valid"] and attempt <= retry:
        if log:
            log("[重试 %d/%d] 上行测量全 INV/NAV（Reliability=%s：%s）。"
                "切换信道后 UE 可能正在重同步/尚未被调度上行发射，"
                "%.1fs 后重新触发测量。" %
                (attempt, retry, d["reliability"], d["reliability_text"], gap))
        if gap:
            time.sleep(gap)
        # 重发本场景的功率命令并等待稳定：切信道/灵敏度段后 UE 可能掉档，
        # 仅重新 INIT 不足以恢复；high 重发 MAXPower，low 重发 CLOop/MINPower。
        set_sens_power(scpi, cfg, power_mode, log)
        d = one_shot()
        attempt += 1

    d["attempts"] = attempt
    d["diagnosis"] = diagnose_reliability(d["reliability"], cfg)
    # NCAP 大片出现时覆盖诊断：Reliability=3 并非真的输入过强，而是 UE 没有
    # 上行发射（RMC 与小区带宽失配 / 未被调度）。给出可执行的排查方向。
    if not d["valid"] and d.get("ncap_ratio", 0.0) > 0.5:
        d["diagnosis"] = (
            "大量子帧 NCAP（未捕获，占比 %.0f%%）→ UE 根本没有上行发射。\n"
            "最可能：RMC/UDCH 的 RB 配置与当前小区带宽失配（例如初始化按 10MHz\n"
            "建小区、RMC 仍 N50，组合却切到 1.4MHz）。请重新点【初始化仪表】\n"
            "（初始化带宽须与待测组合带宽一致）后再测。" % (d["ncap_ratio"] * 100))
    # Reliability=3(Overdriven) 但关键字段有有效读数 → 输入临界过载，实测值
    # （功率/EVM/ACLR）仍可用。记录警告后接受结果；仅关键字段全 INV/NAV 才重试/报错。
    d["overdriven_but_usable"] = bool(d["reliability"] == 3 and d["valid"])
    if d["overdriven_but_usable"] and log:
        log("[警告] MEV Reliability=3 Overdriven（输入临界过载），但功率/关键字段有效，"
            "已接受本次结果：%s" % d["diagnosis"])
    if not d["valid"] and log:
        log("[错误] MEV 结果经 %d 次测量仍全部无效（INV/NAV）。"
            "Reliability=%s：%s\n%s" %
            (attempt, d["reliability"], d["reliability_text"], d["diagnosis"]))
    return d


def sens_min_total_for(cfg, power_mode="high"):
    """按功率档返回 EBLer TotalBlocks(%%) 有效性门限。

    low 档 CLOop -20dBm 上行反馈弱，用 sens_min_total_low（默认 20%）避免
    把真实有效点判为无效；high 档仍用 sens_min_total（默认 50%）。
    """
    low = (power_mode == "low")
    key = "sens_min_total_low" if low else "sens_min_total"
    default = 20.0 if low else 50.0
    raw = cfg.get(key, default)
    if raw is None or str(raw).strip() == "":
        return default
    try:
        return max(0.0, min(100.0, float(raw)))
    except (TypeError, ValueError):
        return default


def measure_sensitivity(scpi, cfg, enabled=None, log=None, band="", bw="", hotspot=None,
                        theory_seed=None, power_mode="high"):
    """灵敏度探底：自适应步进扫描，找满足 Throughput ≥ 门限的最低下行电平。

    判定量 = FETC:LTE:SIGN:EBLer:REL? 的第 2 值 Total Throughput(%)（idx 1，
    2026-11-25 实测对齐参考工具与协议，旧 idx4 判定量偏乐观已修正）；
    Throughput ≥ sens_throughput（默认 95%）判 PASS（与 ASRrun 参考工具一致），
    找仍满足该条件的最低 RS EPRE 电平。
    单点单次 FETC 查询，不再额外查 RELative?，避免响应错位风险并加快扫描。

    仪器设置/扫描的是 RS EPRE (dBm/15kHz)；报告换算为 Full Cell BW Pow (dBm)
    （CMW 手册 23.3.6 / PDF p.1708：P_full = P_rs + 10log10(12*N_RB)）。

    扫描策略（三级）：
      - 强信号区大步进 coarse 快速下降；首次越拐点(FAIL)回退到最后 PASS 点，
        切中步 near；再越拐点回退切精步 step(默认 0.1dBm)，精度到 0.1dB。
      - hotspot = 同频段同带宽上一信道的灵敏度(RS EPRE)，则从其上方 margin 起步、
        在 ±span 窗口内以 near 起步、拐点后切 step 精扫，避免每个信道都全窗重扫。
      - theory_seed = 该 (band,bw) 的 3GPP 理论 REFSENS(RS EPRE)，仅首个信道无
        实测 hotspot 时使用：从 seed+margin 起步但走 coarse 相位、下限仍为全局
        stop（不设窄 floor），兼顾"跳过强信号空扫"与"理论偏差时仍能探到底"。

    按当前 (band,bw) 从 3GPP 表取 UL LCRB@RBstart、DL 满 RB，切 UDCH 精确配置；
    测完切回 RMC 供后续 TX 组合使用。
    返回 dict：result(FullCellBW dBm) / result_epre(RS EPRE dBm, 供热点续扫) / table 等。
    """
    if "sensitivity" not in enabled:
        return None
    cmds = build_bler_commands(cfg, float(cfg["sens_start"]))
    def w(cmd):
        scpi.write(cmd)
        if log:
            log("[SCPI] " + cmd)
    # 切 UDCH 并下发 3GPP REFSENS 的 UL/DL RB（切 STYP 会掉线重连）
    rb_cfg = apply_sens_rb(scpi, band, bw, cfg, log, recheck=reconnect_after_styp)
    # 设置 UE 上行发射功率档（大功率 MAXPower / 小功率 CLOop+CLTPower）
    set_sens_power(scpi, cfg, power_mode, log)
    # 静态配置（一次）
    # OCNG ON：让未使用的 RB 填噪声，使 Full Cell BW Power 恒定、DL 总功率
    # 不随调度 RB 数漂移（手册 p.1709~1710）。ASRrun 参考工具灵敏度段显式
    # 下发该命令；本工具原先缺失，导致不同带宽的 RS EPRE 门限口径不一致
    # （1.4MHz 因无空闲 RB 受影响最大，实测偏 ~6dB）。
    w("CONF:LTE:SIGN:DL:OCNG ON")
    w(cmds["subframes"])
    w(cmds["timeout"])
    # 配置校验：SYST:ERR? 清错误队列（对齐参考工具）——UDCH/RB/OCNG/功率
    # 配置如被仪表拒绝（缺选件等）在此处暴露而非测出全 INV 后才排查
    check_scpi_errors(scpi, log, "灵敏度配置(%s/%s/%s)" % (band, bw, power_mode))
    # 热点值提前解析（warmup 跳过判定与扫描起点共用）
    hot = None
    if hotspot is not None:
        try:
            hot = float(hotspot)
        except (TypeError, ValueError):
            hot = None
    # === EBLer warmup ===
    # 切 UDCH + 设好 TPC 后，UE 需要在当前上行功率下稳定 HARQ ACK/NACK 反馈。
    # 真机观察：小功率档（TPC -20dBm）第一次 EBLer INIT 测量 TotalBlocks 只有 80 多
    # （UE 上行反馈太弱 / 尚未稳定），BLER 被虚高到 19%，导致首信道从种子起步到
    # 退回全窗全是 FAIL、结果为空。先在 -85dBm 强信号下 warmup 一次 EBLer（INIT
    # → RDY → FETCh 丢弃），让 CMW + UE 在当前功率档下先完成一次完整测量，
    # 第二次 INIT 就能拿到 TotalBlocks=100 的准确统计。
    # 大功率档一般不需要 warmup，但加一次也无害。
    # 时长优化（2026-11-25）：热点续扫（同频段同带宽上一信道刚测完）+ 大功率档
    # 时 UE/CMW 已在同一配置下稳定测量过，warmup 纯属重复，跳过省 ~2s/信道；
    # 低功率档（上行功率档已切换）仍保留。
    _need_warmup = not (hot is not None and power_mode == "high")
    _warmup_level = float(cfg["sens_start"])  # -85dBm 强信号
    _warmup_resp = ""
    if _need_warmup:
        w("CONF:LTE:SIGN:DL:RSEPre:LEVel %.1f" % _warmup_level)
        w(cmds["init"])
        _t0 = time.time()
        while time.time() - _t0 < float(cfg.get("bler_timeout", 10)) + 2:
            _st = scpi.query(cmds["state_q"], timeout=2, expect=RE_STAT)
            if "RDY" in _st.upper():
                break
            time.sleep(0.3)
        _warmup_resp = scpi.query(cmds["fetch"], timeout=3, expect=RE_VALUE)
        if log:
            log("[EBLer warmup] 小功率档预同步 1 次（丢弃），FETC=%s" % _warmup_resp[:60])
    elif log:
        log("[EBLer warmup] 热点续扫+大功率档：UE 已稳定，跳过预同步（省 ~2s）")

    def fnum(key, default):
        try:
            return abs(float(cfg.get(key, default))) or default
        except (TypeError, ValueError):
            return default
    start = float(cfg["sens_start"])
    stop = float(cfg["sens_stop"])
    coarse = fnum("sens_coarse", 5.0)
    near = fnum("sens_near", 1.0)
    fine = fnum("sens_step", 0.1)
    thr = float(cfg.get("sens_throughput", 95))  # Throughput 门限(%)，PASS 条件 TPut ≥ thr
    # 统计有效性 = REL? 响应含 BLER 值（idx3）：无块计数（"0,0,,,,"）时 TPut
    # 不可信，该点自动重测一次；仍无效则作为弱侧边界参与定界，不再往更弱
    # 方向下探（真机 124004.log 已证实下探会把 UE 探掉线）。
    # （旧 sens_min_total 按 TotalBlocks 百分比判有效——2026-11-25 核实 REL?
    # idx1 实为吞吐量而非块数百分比，该机制语义有误已移除；配置键保留兼容。）
    # UE 掉线特征门限：TPut 低于该值时再查 PSW:STAT 确认是否真脱网；
    # 仅 TPut 低不足以判掉线（真机 145949.log 证实会误杀 bisect、结果虚好）。
    try:
        drop_total = float(cfg.get("sens_drop_total", 10) or 10)
    except (TypeError, ValueError):
        drop_total = 10.0
    drop_total = max(0.0, min(drop_total, 100.0))
    # 全局参数 + 热点独立参数（向后兼容：config.json 没配 hotspot_coarse/seed_margin 时
    # 分别回退到 sens_coarse 和 sens_hotspot_margin 的旧语义）
    hs_margin = fnum("sens_hotspot_margin", 3.0)
    hs_span = fnum("sens_hotspot_span", 3.0)
    hs_coarse = fnum("sens_hotspot_coarse", 2.0) or coarse
    # seed_margin 独立参数，未显式配置时沿用 hs_margin（向后兼容：旧版本只用一个 margin
    # 同时服务种子起步和热点续扫）；用户显式配 sens_seed_margin 时才生效。
    _seed_raw = cfg.get("sens_seed_margin")
    if _seed_raw is None or str(_seed_raw).strip() == "":
        seed_margin = hs_margin
    else:
        seed_margin = fnum("sens_seed_margin", hs_margin)
    if log:
        log("[灵敏度判定] 标准：Throughput >= %.1f%%（FETC:LTE:SIGN:EBLer:REL? "
            "第2值 Total TPut，%s 档）" % (thr, power_mode))

    ascending = stop > start
    # 扫描中是否确已把 UE 探到脱网（PSW=ON）。收尾据此决定是否在强信号下
    # 耐心等待自恢复；未脱网的组合零开销跳过（避免拖慢正常组合）。
    ue_dropped = False
    seed = None
    if theory_seed is not None:
        try:
            seed = float(theory_seed)
        except (TypeError, ValueError):
            seed = None
    # steps 字典构建：热点续扫覆盖 coarse 步长为更小的 hs_coarse，
    # 种子起步用全局 coarse，全窗扫描也用全局 coarse。
    if hot is not None and ascending is False:
        # 实测热点（同频段上一信道）：已有先验值，用 hs_coarse 步长快速逼近拐点，
        # 首次越拐点即触发二分精扫（不再走 near→fine 线性扫）。
        # 典型窗口宽度 = hs_margin + hs_span ≈ 6dB，hs_coarse=2dB 约 3-4 步到位。
        level = max(stop, min(start, hot + hs_margin))
        floor = max(stop, hot - hs_span)
        phase = "coarse"
        steps = {"coarse": hs_coarse, "near": near, "fine": fine}
        if log:
            log("[灵敏度扫描] %s/%s 热点续扫：起点 %.1f (上次 %.1f +%.0f)，"
                "续扫窗口 [%.1f, %.1f] dBm RS EPRE（粗步 %.1f → 拐点后 0.1dB 二分精扫）" %
                (band_label(band), bw_label(bw), level, hot, hs_margin,
                 floor, level, hs_coarse))
    elif seed is not None and ascending is False:
        # 理论种子（首信道）：从种子值+seed_margin 起步，走 coarse 相位、下限仍为全局
        # stop，不设窄 floor。跳过 -85 起的强信号空扫（省约 13s），理论值与实际偏差
        # 较大时 coarse 大步进仍能快速找到拐点后二分。
        level = max(stop, min(start, seed + seed_margin))
        floor = stop
        phase = "coarse"
        steps = {"coarse": coarse, "near": near, "fine": fine}
        if log:
            log("[灵敏度扫描] %s/%s 理论种子起步：起点 %.1f (3GPP REFSENS %.1f +%.0f)，"
                "粗步 %.1f → 拐点后 0.1dB 二分精扫，下限 %.1f dBm RS EPRE" %
                (band_label(band), bw_label(bw), level, seed, seed_margin,
                 coarse, floor))
    else:
        level = start
        floor = stop
        phase = "coarse"
        steps = {"coarse": coarse, "near": near, "fine": fine}
        if log:
            log("[灵敏度扫描] %s/%s 全窗扫描：起点 %.1f，粗步 %.1f → 拐点后 0.1dB 二分精扫，"
                "下限 %.1f dBm RS EPRE" %
                (band_label(band), bw_label(bw), level, coarse, floor))

    table = []
    result = None            # 最后满足门限的电平（RS EPRE）
    last_pass = None
    seen = set()             # 已探测电平（二分时避免重复点）
    lo = hi = None           # 二分窗口：lo=FAIL/无效(弱侧), hi=PASS(强侧)
    weak_floor = None        # 无效点弱边界：精扫只在 >该电平一侧进行，绝不继续下探
    bisect_ok = not ascending  # 下降扫描启用二分精扫；上升扫描沿用原线性
    guard = 0
    while guard < 200:
        guard += 1
        w("CONF:LTE:SIGN:DL:RSEPre:LEVel %.1f" % level)
        w("INIT:LTE:SIGN:EBLer")
        ok, st = wait_for(scpi, cmds["state_q"],
                          lambda s: "RDY" in s.upper(),
                          float(cfg["bler_timeout"]) + 3.0, 0.3, log,
                          expect=RE_STAT)
        resp = scpi.query(cmds["fetch"], expect=RE_VALUE)
        if log:
            log("[SCPI] %s -> %s" % (cmds["fetch"], resp))
        vals = parse_values(resp)
        # FETC:EBLer:REL? 返回 6 值（2026-11-25 对齐 ASRrun 参考工具与协议实测确认）：
        #   Rel(0), TotalTPut%(1), DTX%(2), BLER%(3), 副统计(4), 副统计(5)
        # idx3 = 100 - idx1 恒成立（BLER+TPut=100）。协议判定量（TS 36.521
        # DL-SCH 总吞吐 ≥95%）= idx1：参考工具组合1 -125.3 响应 idx1=95.556
        # 与其数据表 TPut=95.56 完全吻合；旧代码取 idx4 是另一个略偏乐观的量
        # （-125.4 处 idx1=94.26 应 FAIL 而 idx4=95.22 假 PASS），本日修正。
        tput = at(vals, 1)   # Total Throughput(%)：协议判定量
        bler = at(vals, 3)   # BLER(%)：有统计块时必为数值，缺失=统计无效
        # 统计无效（响应形如 "0,0,,,,"，无块计数 → TPut 不可信）自动重测一次，
        # 仍无效才判该点为无效点（弱边界）。
        if bler is None:
            if log:
                log("[灵敏度扫描] %.1f dBm 无统计块（BLER 缺失），重测一次" % level)
            w("INIT:LTE:SIGN:EBLer")
            ok, st = wait_for(scpi, cmds["state_q"],
                              lambda s: "RDY" in s.upper(),
                              float(cfg["bler_timeout"]) + 3.0, 0.3, log,
                              expect=RE_STAT)
            resp = scpi.query(cmds["fetch"], expect=RE_VALUE)
            if log:
                log("[SCPI] %s -> %s" % (cmds["fetch"], resp))
            vals = parse_values(resp)
            tput = at(vals, 1)
            bler = at(vals, 3)
        sample_ok = (bler is not None)
        passed = sample_ok and (tput is not None and tput >= thr) and ok   # ≥门限(默认 95%)
        table.append({"level": level, "level_bw": epre_to_fullbw(level, bw),
                      "throughput": tput, "bler": bler, "sample_ok": sample_ok,
                      "invalid": not sample_ok,
                      "state": st, "pass": passed,
                      "phase": "bisect" if phase == "bisect" else phase,
                      "raw": resp})
        seen.add(round(level, 3))
        # 方案 A 修正：TotalBlocks 低本身不足以判掉线——极弱信号下 UE 仍在网
        # 但收不到 PDCCH，Total 同样会掉到 0~2%（真机 145949.log 证实：4 次
        # 粗扫 FAIL 点 Total=0~0.17% 被误判掉线，bisect 全被跳过，10MHz 结果
        # 虚好 4.3dB）。必须用 PSW:STAT 确认 UE 确实脱网（ON）才标记 dropped。
        _psw = ""
        _psw_queried = False
        if tput is not None and tput < drop_total:
            try:
                _psw = scpi.query(CONN_RECHECK_QUERY, timeout=2.0, expect=RE_PSW)
                _psw_queried = True
            except Exception:
                _psw = ""
            table[-1]["psw"] = _psw
        if _psw_queried and str(_psw).strip() and not conn_ready(_psw):
            # UE 确实脱网：该点 EBLer 无意义，立即停止，不当弱边界反复二分
            ue_dropped = True
            table[-1]["dropped"] = True
            if log:
                log("[灵敏度扫描] %.1f dBm TPut=%s%% < %.0f%% 且 UE 已脱网"
                    "（%s=%s），判定掉线，停止二分与下探；保留上一有效电平 %s dBm" %
                    (level, fmt(tput, 2) or "--", drop_total, CONN_RECHECK_QUERY,
                     str(_psw).strip(), fmt(result, 1) or "--"))
            break
        if _psw_queried and str(_psw).strip() and conn_ready(_psw):
            # UE 仍在网：TPut 低只是信号弱（真实 FAIL），按普通 FAIL/无效点继续
            # 走下面的定界逻辑，绝不误判掉线（避免假改善）。
            if log:
                log("[灵敏度扫描] %.1f dBm TPut=%s%% < %.0f%% 但 UE 仍在网"
                    "（%s=%s），继续定界" %
                    (level, fmt(tput, 2) or "--", drop_total, CONN_RECHECK_QUERY,
                     str(_psw).strip()))
        if not sample_ok:
            # 统计无效（重测后仍无块计数）：该点数据不可信，既不能当 FAIL 拐点，
            # 也不能继续往更弱方向下探（真机证明会探掉线、后续全盘失败）。
            # 真机 085027.log 组合1（1.4M/低）实证根因：coarse 5dB 大步直接跳进
            # 失步区（-122.6→-127.6）UE 瞬间失步，此后所有电平统计全空，弱边界
            # 层层收窄把结果塌缩在 coarse PASS 点（早停 2.7dB，参考工具同点
            # 1dB 步进全程有效统计、真实拐点 -125.3）。对策：非二分相位的首次
            # 无效点 → 退回上一有效 PASS 电平、改 1dB 步进重降（UE 对渐变信号
            # 始终保持跟踪）；1dB 步进仍命中无效点才作为弱边界进入二分定界。
            if phase not in ("bisect", "near") and last_pass is not None \
                    and bisect_ok:
                if log:
                    log("[灵敏度扫描] %.1f dBm 统计无效（coarse 大步跳进失步区），"
                        "退回上一有效电平 %.1f 改 %.1fdB 步进重降（保持 UE 跟踪）" %
                        (level, last_pass, near))
                level = last_pass
                phase = "near"
                continue
            if phase == "bisect" and lo is not None and hi is not None:
                if log:
                    log("[灵敏度扫描] %.1f dBm 统计无效（无块计数），"
                        "作为二分弱边界，窗口收窄 [%.1f, %.1f]，不再往更弱下探" %
                        (level, level, hi))
                lo = level
                weak_floor = level
            elif last_pass is not None and bisect_ok:
                lo, hi, phase = level, last_pass, "bisect"
                weak_floor = level
                if log:
                    log("[灵敏度扫描] %.1f dBm 统计无效（无块计数），"
                        "作为弱边界启动 0.1dB 精扫 [%.1f, %.1f]（不再往更弱下探；"
                        "上一有效电平 %s dBm）" %
                        (level, lo, hi, fmt(result, 1) or "--"))
            else:
                if log:
                    log("[灵敏度扫描] %.1f dBm 统计无效（无块计数），"
                        "无有效强边界可定界，停止继续下探；保留上一有效电平 %s dBm" %
                        (level, fmt(result, 1) or "--"))
                break
        elif passed:
            result = level
            last_pass = level
            if phase == "bisect":
                hi = level
        else:
            if phase != "bisect" and last_pass is not None and bisect_ok:
                # coarse 首次越拐点：锁定二分窗口 [本次FAIL, 上次PASS]，随后
                # 在 <=coarse 的窗口内二分到 0.1dB 精度（约 log2(coarse/fine)
                # 个点），取代原 near+fine 线性扫描（后者最多 near/fine+10 点）。
                lo, hi, phase = level, last_pass, "bisect"
                if log:
                    log("[灵敏度扫描] 拐点出现于 %.1f dBm，二分精扫窗口 [%.1f, %.1f] "
                        "（0.1dB 精度）" % (level, lo, hi))
            elif phase == "bisect":
                lo = level
            elif level != start and not ascending:
                # 起点即 FAIL（理论种子偏高 / UE 灵敏度远差于理论）：退回全窗起点重扫
                if log:
                    log("[灵敏度扫描] 起点 %.1f dBm 即 FAIL，退回全窗起点 %.1f 重扫" %
                        (level, start))
                level = start
                continue
            else:
                break
        if phase == "bisect":
            if lo is None or hi is None or (hi - lo) <= fine + 1e-9:
                break
            # 窗口过宽（典型 coarse=5dB 触发 bisect 时 5dB 宽）：首次 bisect 迭代
            # 不在中点，而是从 FAIL 侧往里 hs_coarse 步缩窗，避免首个 mid 直接
            # 打到 FAIL 区浪费 1 次测量（后续 bisect 自动收敛到 0.1dB）。
            _first_bisect = not any(r.get("phase") == "bisect" for r in table)
            if _first_bisect and (hi - lo) > 2 * hs_coarse:
                mid = round(lo + hs_coarse, 1)
            else:
                mid = round((lo + hi) / 2.0, 1)
            if mid in seen or mid <= lo or mid >= hi:
                # 中点已测或落到边界：在 (lo,hi) 内找最近的 0.1 网格未测点
                cand = round(lo + fine, 1)
                while cand < hi and (cand in seen or cand <= lo):
                    cand = round(cand + fine, 1)
                if cand >= hi or cand in seen:
                    break
                mid = cand
            # 弱边界硬约束：任何情况下都不再探测 <= 无效点的电平（防 UE 掉线）
            if weak_floor is not None and mid <= weak_floor:
                break
            level = mid
        else:
            nxt = level + steps[phase] if ascending else level - steps[phase]
            if (ascending and nxt > floor) or (not ascending and nxt < floor):
                # 最后一步越过 floor：若尚未测到 floor，则钳到 floor 再测一次，
                # 否则极端偏差（UE 灵敏度远差于理论）时拐点夹在最后一点与 floor
                # 之间，不测 floor 就 break 会漏探到底。
                if abs(level - floor) < 1e-6:
                    break
                nxt = floor
            level = round(nxt, 3)   # 0.1 步进避免浮点累积漂移
    # === 扫描收尾：先拉回强信号，再决定是否需要切回 RMC / 重附着 ===
    # 真机 151333.log 证实旧顺序有严重耗时问题：扫描停在 -123dBm 弱信号时
    # 先下发 STYP RMC，把仍在网 UE 搞掉线，随后 ACT CONN 5s 失败→重建小区，
    # 组合 1/3 各白耗 ~15s。用户确认的正确流程：同带宽测完一个信道后，
    # 先把 RS EPRE 恢复到初始强信号，直接切下一信道再调电平测，无需重走注册。
    try:
        _strong = float(cfg.get("sens_start", -85))
    except (TypeError, ValueError):
        _strong = -85.0
    w("CONF:LTE:SIGN:DL:RSEPre:LEVel %.1f" % _strong)
    if log:
        log("[灵敏度收尾] RS EPRE 先恢复到初始强信号 %.1f dBm"
            "（供下一信道直接切换，避免弱信号下切 STYP 掉线）" % _strong)
    time.sleep(0.5)  # 让仪表电平生效、UE 测量到强信号后再判断注册态

    # 扫描把 UE 探掉线时（152654.log 组合 3 高信道）：UE 从极弱信号失步后，
    # 强信号下需要数秒重新同步（真机重建后 6s 内自行 ATT）。旧逻辑仅 sleep 0.5s
    # 就查一次 PSW，恰逢仪表忙超时 → 空状态被当掉线 → 立即 ACT CONN 5s 全 ON
    # → 白耗 18s 重建小区。此处改为在强信号下耐心轮询等 UE 自恢复，wait_for
    # 容忍单次查询超时；恢复即跳过重附着/重建，未恢复才走下面的 reconnect。
    if ue_dropped:
        try:
            _recov = max(1.0, min(15.0, float(cfg.get("attach_settle", 2) or 2) * 3))
        except (TypeError, ValueError):
            _recov = 6.0
        if log:
            log("[灵敏度收尾] 扫描中 UE 已脱网，强信号下等待最多 %.1fs 让其自行重同步"
                "（避免查询超时误判后直接重建小区）" % _recov)
        _ok_self, _st_self = wait_for(scpi, CONN_RECHECK_QUERY, conn_ready,
                                      _recov, 0.5, log, query_timeout=2.0,
                                      expect=RE_PSW)
        if _ok_self:
            if log:
                log("[灵敏度收尾] UE 已在强信号下自行恢复注册（%s），跳过重附着/重建小区" %
                    _st_self)
            ue_dropped = False   # 已恢复，下面 reconnect 会零开销返回

    # 只有本轮还勾选了 TX 指标时，后续 TX 组合才需要 RMC 满 RB 连接：
    # 此时已在强信号下切 RMC，不会把 UE 搞掉线（弱信号切 STYP 才是根因）。
    # 纯灵敏度场景完全不切 RMC——下一信道的 apply_sens_rb 会自行重配，
    # 同带宽换信道 UE 本就保持注册（151333.log 组合 2/5/6 已证实）。
    need_rf = bool(enabled & {"maxpower", "aclr", "evm", "freqerr"})
    if need_rf:
        for cmd in restore_rmc_commands(cfg, bw):
            w(cmd)
        if log:
            log("[灵敏度RB] 已恢复 RMC 连接（本轮含 TX 指标，强信号下切换）")

    # 轻量注册复查：强信号下若 UE 仍在网则零开销返回（不重附着、不重建小区）；
    # 仅当扫描确已把 UE 探掉线时，才在强信号下 ACT CONN 秒级恢复。
    if reconnect_after_styp:
        try:
            reconnect_after_styp(scpi, log, cfg=cfg,
                                 retune_ctx={"band": band, "bw": bw, "earfcn": None})
        except Exception as exc:
            raise ScpiError(
                "灵敏度测完后 UE 在强信号下仍无法恢复注册（含快速重建小区）：%s。"
                "已中止以避免后续无效数据。" % exc)
    result_bw = epre_to_fullbw(result, bw)
    if log:
        log("[灵敏度扫描] 完成：%d 点，RS EPRE=%s dBm，Full Cell BW Pow=%s dBm，"
            "判定标准 Throughput >= %.1f%%" %
            (len(table), fmt(result, 1) or "--", fmt(result_bw, 2) or "--", thr))
    return {"result": result_bw, "result_epre": result, "threshold": thr,
            "criterion": "Throughput >= %.1f%%" % thr,
            "level_label": "Full Cell BW Pow (dBm)", "table": table,
            "power_mode": power_mode,
            "rb_config": "%s/%s UL %d@%d DL满RB%d [%s]" %
                         (band_label(band), bw_label(bw),
                          rb_cfg.get("rb_applied", rb_cfg["rb"])[0],
                          rb_cfg.get("rb_applied", rb_cfg["rb"])[1], rb_cfg["dl_rb"],
                          rb_cfg.get("mode", "RMC")),
            "raw": "\n".join("RS %s / FullBW %s dBm | TPut=%s%% | %s" %
                             (fmt(r["level"], 1), fmt(r["level_bw"], 2),
                              fmt(r["throughput"]), "PASS" if r["pass"] else "FAIL")
                             for r in table)}


# ----------------------------------------------------------------------------
# 结果整理：把测量数据变成表格行
# ----------------------------------------------------------------------------


def make_rows(data, limits=None):
    """把单次组合的测量 dict 整理为 [(类别, 指标值, 单位, 判定), ...]。"""
    limits = limits or {}
    rows = []
    ma = data.get("mod_avg") or []
    mx = data.get("mod_ext") or []
    ac = data.get("aclr") or []

    def judge_max(v, lim):
        if v is None or lim in (None, ""):
            return ""
        try:
            return "PASS" if v <= float(lim) else "FAIL"
        except (TypeError, ValueError):
            return ""

    def judge_min(v, lim):
        if v is None or lim in (None, ""):
            return ""
        try:
            return "PASS" if v >= float(lim) else "FAIL"
        except (TypeError, ValueError):
            return ""

    # 索引沿用手册 21.5.3.26 调制单值返回结构
    tx_power = at(ma, 17)
    ac_uttra2_lo = at(ac, 1)
    ac_utra1_lo = at(ac, 2)
    ac_eutra_lo = at(ac, 3)
    ac_eutra_hi = at(ac, 5)
    ac_utra1_hi = at(ac, 6)
    ac_utra2_hi = at(ac, 7)
    evm_rms = at(ma, 2)
    evm_peak = at(ma, 4)
    freq_err = at(ma, 15)

    rows.append(["功率", "Power (TX)", fmt(tx_power), "dBm",
                 judge_max(tx_power, limits.get("evm_lim")) if False else ""])
    rows.append(["ACLR", "E-UTRA 下邻道", fmt(ac_eutra_lo), "dB",
                 judge_min(ac_eutra_lo, limits.get("aclr_lim"))])
    rows.append(["ACLR", "E-UTRA 上邻道", fmt(ac_eutra_hi), "dB",
                 judge_min(ac_eutra_hi, limits.get("aclr_lim"))])
    rows.append(["ACLR", "UTRA1 下邻道", fmt(ac_utra1_lo), "dB", ""])
    rows.append(["ACLR", "UTRA1 上邻道", fmt(ac_utra1_hi), "dB", ""])
    rows.append(["ACLR", "UTRA2 下邻道", fmt(ac_uttra2_lo), "dB", ""])
    rows.append(["ACLR", "UTRA2 上邻道", fmt(ac_utra2_hi), "dB", ""])
    rows.append(["EVM", "EVM RMS (low)", fmt(evm_rms), "%",
                 judge_max(evm_rms, limits.get("evm_lim"))])
    rows.append(["EVM", "EVM Peak (low)", fmt(evm_peak), "%", ""])
    rows.append(["频率误差", "载波频率误差", fmt(freq_err), "Hz",
                 judge_max(abs(freq_err) if freq_err is not None else None,
                           limits.get("freq_lim"))])

    psu = data.get("psu") or {}
    if psu:
        rows.append(["电流", "对应电流 (发射)", fmt(psu.get("current"), 1), "mA", ""])
        if psu.get("voltage") is not None:
            rows.append(["电流", "电源电压", fmt(psu.get("voltage"), 3), "V", ""])
        if psu.get("power") is not None:
            rows.append(["电流", "电源功率", fmt(psu.get("power"), 3), "W", ""])

    sens = data.get("sensitivity") or {}
    if sens:
        rows.append(["灵敏度", "灵敏度 (Full Cell BW)", fmt(sens.get("result"), 2), "dBm",
                     judge_max(sens.get("result"), limits.get("sens_lim"))])
    return rows


# ----------------------------------------------------------------------------
# 宽表：一行 = 一个测试组合（频段/带宽/信道），一列 = 一个指标
# ----------------------------------------------------------------------------

# (键, 表头, 单位, 判定方向)  判定方向: max=越大越好/不超过, min=不低于, none=不判定
# 完整列顺序：Band/BW/Channel → Voltage → Current(拆分@Pmax/@Pmin) → Pmax → ACLR → EVM → FreqErr → LCRB @ RBstart → Sens → Desense
# 列名已统一英文；纯 TX 测试（无 RX 灵敏度）时 LCRB @ RBstart/Sens limit/Sens@Pmax/Sens@Pmin/Desense 自动隐藏。
WIDE_COLUMNS = [
    ("voltage", "Voltage\n(V)", "", None),
    ("current", "Current\n(A)", "", None),
    ("maxpower", "Power\n(dBm)", "", None),
    ("aclr_eutra_lo", "E-UTRA1-L\n(dB)", "", "min"),
    ("aclr_eutra_hi", "E-UTRA1-R\n(dB)", "", "min"),
    ("aclr_utra1_lo", "UTRA1-L\n(dB)", "", None),
    ("aclr_utra1_hi", "UTRA1-R\n(dB)", "", None),
    ("aclr_utra2_lo", "UTRA2-L\n(dB)", "", None),
    ("aclr_utra2_hi", "UTRA2-R\n(dB)", "", None),
    ("evm_rms", "EVM RMS\n(%)", "", "max"),
    ("freqerr", "Freq Error\n(Hz)", "", "max"),
    ("sensitivity", "Sens\n(dBm)", "", None),
]
# 判定相关列已按需求去掉：结果表只保留数值列（宽表），不再显示 PASS/FAIL 判定列。
# 置空后 wide_headers / make_wide_row / export / GUI 表格均自动退化为"无判定列"，
# 相关 len() 计算仍自洽（无需改动其它函数）。
VERDICT_COLUMNS = []


def _col_metric(key):
    """宽表列 key → 所属可勾选指标键；电源/组合列返回 None（始终显示）。"""
    if key in ("voltage", "current_pmax", "current_pmin", "current_pa",
               "ul_rb", "sens_limit", "desense"):
        return None
    if key == "maxpower":
        return "maxpower"
    if key.startswith("aclr"):
        return "aclr"
    if key == "evm_rms":
        return "evm"
    if key == "freqerr":
        return "freqerr"
    if key.startswith("sensitivity"):
        return "sensitivity"
    return None


def _build_fail_flags(wide_values, combo_label, cfg, headers=None):
    """返回一个与 wide_values 等长的 bool 列表：True=该单元格 FAIL（需标红），False=正常。

    用每列独立 FAIL flag 代替 Treeview 的行级 tag，实现单元格级标红。
    组合列（频段/带宽/信道[+/UL RB/sens_limit]）永远 False；
    指标列按 active_columns 给的方向判定。
    """
    try:
        cols = active_columns(
            cfg.get("metrics"),
            cfg.get("sens_power_mode", "high"),
            bool(cfg.get("psu_enable", True)),
        )
    except Exception:
        return [False] * len(wide_values)
    try:
        lims = {
            "evm_lim": cfg.get("evm_lim", ""),
            "aclr_lim": cfg.get("aclr_lim", ""),
            "freq_lim": cfg.get("freq_lim", ""),
            "sens_lim": cfg.get("sens_lim", ""),
        }
    except Exception:
        lims = {}
    # combo 列固定 3 列（band/bw/chan）；指标列全由 active_columns 管理，
    # UL RB / sens_limit 无判定方向（direction=None），循环里自然跳过。
    prefix = 3
    flags = [False] * len(wide_values)
    for i, col_def in enumerate(cols):
        idx = prefix + i
        if idx >= len(wide_values):
            break
        col_key = col_def[0]
        val = wide_values[idx]
        try:
            fv = float(val)
        except (TypeError, ValueError):
            continue
        fail = False
        if col_key.startswith("evm"):
            lim = lims.get("evm_lim", "")
            if lim:
                try: fail = abs(fv) > float(lim)
                except ValueError: pass
        elif col_key.startswith("freqerr"):
            lim = lims.get("freq_lim", "")
            if lim:
                try: fail = abs(fv) > float(lim)
                except ValueError: pass
        elif col_key.startswith("aclr"):
            lim = lims.get("aclr_lim", "")
            if lim:
                try: fail = fv < float(lim)
                except ValueError: pass
        elif col_key.startswith("sensitivity"):
            lim = lims.get("sens_lim", "")
            if lim:
                try: fail = fv > float(lim)
                except ValueError: pass
        flags[idx] = fail
    return flags


def _wide_row_tag(wide_values, combo_label, cfg):
    """兼容旧调用：整行任一 FAIL → 'fail'，否则 ''。已弃用，用 _build_fail_flags。"""
    flags = _build_fail_flags(wide_values, combo_label, cfg)
    return "fail" if any(flags) else ""


def _scenario_tag(mode):
    """测试场景 → 列名后缀（TX/RX 公用场景）。"""
    return {"high": "@Pmax", "low": "@Pmin"}.get(mode, "")


def _tx_col_label(label, mode):
    """给 TX 指标表头加场景标注：'Power\\n(dBm)' → 'Power@Pmax\\n(dBm)'。"""
    tag = _scenario_tag(mode)
    if not tag:
        return label
    if "\n" in label:
        head, unit = label.split("\n", 1)
        return head + tag + "\n" + unit
    return label + tag


# 形如 "Power@Pmax" / "Sens@Pmin" 的场景后缀（只认 Pmax/Pmin，
# 避免误伤 "LCRB @ RBstart" 这类本身含 @ 的正常列名）
_SCEN_SUFFIX_RE = re.compile(r"^(.*[^@\s])@(Pmax|Pmin)$")


def header_display(label):
    """界面表头显示用：名称占第一行，场景标注 + 单位占第二行（共 2 行）。

    只影响 GUI 渲染，内部列名/导出表头/自检断言仍用原字符串，
    避免 txt 按行拼接时的换行污染。
      'Power@Pmax\\n(dBm)'   → 'Power\\n@Pmax (dBm)'
      'Current@Pmin\\n(mA)'  → 'Current\\n@Pmin (mA)'
      'Sens@Pmax\\n(dBm)'    → 'Sens\\n@Pmax (dBm)'
      'Power@Pmax'           → 'Power\\n@Pmax'
      'LCRB @ RBstart'       → 原样返回
    """
    if not label or "@" not in label:
        return label
    head, _, unit = label.partition("\n")
    m = _SCEN_SUFFIX_RE.match(head)
    if not m:
        return label
    out = m.group(1) + "\n" + "@" + m.group(2)
    if unit:
        out += " " + unit
    return out


def active_columns(metrics=None, sens_power_mode="high", psu_enable=True):
    """按当前勾选指标 + 灵敏度功率档，返回应显示的宽表列定义。

    - 未勾选的指标列不显示（电压/电流随电源，不属于勾选项）。
    - 电压/电流是 TX 测量时随电源记录的功耗：只勾 RX 灵敏度（无任何 TX 指标）时不显示。
    - UL RB + sens Limit 仅在勾选 sensitivity 时显示，排在 TX 指标之后、灵敏度列之前。
    - 灵敏度：both 显示「sens大/sens小」两列 + Desense；low 显示单列；high 显示单列。
    - 电流：both 模式下拆成「电流(大)」+「电流(小)」两列；其他模式保持单列。
    """
    if metrics is None:
        metrics = {k: True for k in METRIC_KEYS}
    has_tx = any(metrics.get(k) for k in TX_METRIC_KEYS)
    has_sens = bool(metrics.get("sensitivity"))
    mode = sens_power_mode or "high"
    cols = []
    for c in WIDE_COLUMNS:
        key = c[0]
        if key in ("voltage", "current") and (not psu_enable or not has_tx):
            continue
        m = _col_metric(key)
        if m is not None and not metrics.get(m, False):
            continue
        if key == "voltage":
            cols.append(("voltage", "Voltage\n(V)", "", None))
            continue
        if key == "current":
            # 三种模式列名固定 Current@Pmax / Current@Pmin / Current of PA，
            # 根据 RX 功率档显隐：
            #   high → Current@Pmax（Pmax 档电流）
            #   low  → Current@Pmin（Pmin 档电流）
            #   both → Current@Pmax + Current@Pmin + Current of PA
            #   注：其他 TX 指标（Power/ACLR/EVM/频偏）跟随 RX 功率档取对应值
            if mode == "high":
                cols.append(("current_pmax", "Current@Pmax\n(mA)", "", None))
            elif mode == "low":
                cols.append(("current_pmin", "Current@Pmin\n(mA)", "", None))
            else:  # both
                cols.append(("current_pmax", "Current@Pmax\n(mA)", "", None))
                cols.append(("current_pmin", "Current@Pmin\n(mA)", "", None))
                cols.append(("current_pa", "Current of PA\n(mA)", "", None))
            continue
        if m in TX_METRIC_KEYS:
            # TX 指标按公用测试场景出列：both 拆成 @Pmax/@Pmin 双列，
            # 单场景只出一列但列名同样带档位标注，避免误读。
            if mode == "both":
                cols.append((key + "_high", _tx_col_label(c[1], "high"), c[2], c[3]))
                cols.append((key + "_low", _tx_col_label(c[1], "low"), c[2], c[3]))
            else:
                cols.append((key, _tx_col_label(c[1], mode), c[2], c[3]))
            continue
        if key == "sensitivity":
            # UL RB + sens Limit 插到灵敏度列之前（仅当勾选灵敏度）
            if has_sens:
                cols.append(("ul_rb", "LCRB @ RBstart", "", None))
                cols.append(("sens_limit", "Sens limit\n(dBm)", "", None))
            if has_sens:
                if mode == "both":
                    cols.append(("sensitivity_high", "Sens@Pmax\n(dBm)", "", None))
                    cols.append(("sensitivity_low", "Sens@Pmin\n(dBm)", "", None))
                    cols.append(("desense", "Desense\n(dB)", "", None))
                elif mode == "low":
                    cols.append(("sensitivity_low", "Sens@Pmin\n(dBm)", "", None))
                else:
                    cols.append(("sensitivity", "Sens@Pmax\n(dBm)", "", None))
            continue
        cols.append(c)
    return cols


def wide_headers(metrics=None, sens_power_mode="high", psu_enable=True):
    """宽表表头：[Band, BW, Channel] + active_columns 输出的指标列。"""
    cols = active_columns(metrics, sens_power_mode, psu_enable)
    # 本工具实际写入的是下行信道 EARFCN（CONF:LTE:SIGN:RFS:CHAN:DL），
    # 信道表也按 DL EARFCN 组织；TDD 频段 UL/DL 同频，仍按 DL 口径标记。
    combo_cols = ["Band", "BW", "DL Channel"]
    return combo_cols + [c[1] for c in cols] + [c[1] for c in VERDICT_COLUMNS]


def make_wide_row(combo_label, data, limits=None, metrics=None,
                  sens_power_mode="high", psu_enable=True):
    """把一个组合的测量数据整理成宽表一行（与 wide_headers 对齐）。

    combo_label 形如 "B1 / 10MHz / 低(CH50)"，拆成 频段/带宽/信道 三列。
    只输出当前勾选指标对应的列；判定列已去掉，limits 仅为兼容旧调用保留。
    data["sensitivity"] 为大功率结果，data["sensitivity_low"] 为小功率结果。
    """
    ma = data.get("mod_avg") or []
    ac = data.get("aclr") or []
    psu = data.get("psu") or {}
    psu_hi = data.get("psu_high") or {}
    psu_lo = data.get("psu_low") or {}
    sens_hi = data.get("sensitivity") or {}
    sens_lo = data.get("sensitivity_low") or {}

    parts = [p.strip() for p in combo_label.split("/")]
    band = parts[0] if len(parts) > 0 else combo_label
    bw = parts[1] if len(parts) > 1 else ""
    chan = parts[2] if len(parts) > 2 else ""
    # 结果表 DL Channel 列只显示信道号（CH607），去掉 低/中/高/全 前缀与括号。
    # 组合标签本身（含位置）仍用于日志/限值解析，不在这里改动。
    _m = re.search(r"CH\s*(\d+)", chan)
    if _m:
        chan = "CH" + _m.group(1)

    mode = sens_power_mode or "high"
    # TX 指标按公用测试场景取值：both 同时输出 @Pmax/@Pmin 两套；
    # 单场景只输出该档一套（缺该档数据时兜底用另一档，兼容旧报告数据）。
    def _tx_vals(_ma, _ac):
        return {
            "maxpower": fmt(at(_ma, 17)),
            "aclr_eutra_lo": fmt(at(_ac, 3)),
            "aclr_eutra_hi": fmt(at(_ac, 5)),
            "aclr_utra1_lo": fmt(at(_ac, 2)),
            "aclr_utra1_hi": fmt(at(_ac, 6)),
            "aclr_utra2_lo": fmt(at(_ac, 1)),
            "aclr_utra2_hi": fmt(at(_ac, 7)),
            "evm_rms": fmt(at(_ma, 2)),
            "freqerr": fmt(at(_ma, 15)),
        }
    tx_hi = _tx_vals(ma, ac)
    # Pmin 档 TX：必须是「真的在 Pmin 档测过」的数据（mod_avg_low/aclr_low）。
    # both 场景缺这组数据时留空——绝不回退复制 Pmax 的值，否则 @Pmin 列看起来
    # 有数、实际是 Pmax 的重复（171835.log 暴露的问题）。单档 low 场景保留
    # 回退（外部通常已把该档结果同时写进 mod_avg/mod_avg_low）。
    if mode == "low":
        tx_lo = _tx_vals(data.get("mod_avg_low") or ma,
                         data.get("aclr_low") or ac)
    else:
        tx_lo = _tx_vals(data.get("mod_avg_low"), data.get("aclr_low"))
    # 电流列按场景显隐：high→@Pmax，low→@Pmin，both→两列 + PA 差值
    vals = {
        "voltage": fmt((psu_hi or psu_lo or psu).get("voltage"), 3),
        "current_pmax": fmt((psu_hi or psu).get("current"), 1),
        "current_pmin": fmt((psu_lo or psu).get("current"), 1),
        # Current of PA = Current@Pmax - Current@Pmin（PA 自身静态工作电流，mA）
        "current_pa": fmt(((psu_hi or psu).get("current") - (psu_lo or psu).get("current"))
                           if ((psu_hi or psu).get("current") is not None and
                               (psu_lo or psu).get("current") is not None) else None, 1),
    }
    vals.update(tx_lo if mode == "low" else tx_hi)
    vals.update({k + "_high": v for k, v in tx_hi.items()})
    vals.update({k + "_low": v for k, v in tx_lo.items()})
    vals.update({
        "sensitivity": fmt(sens_hi.get("result"), 1),
        "sensitivity_high": fmt(sens_hi.get("result"), 1),
        "sensitivity_low": fmt(sens_lo.get("result"), 1),
        # desense = 大功率灵敏度 - 小功率灵敏度（两者均为 FullBW dBm，差值即 dB）
        "desense": fmt((sens_hi.get("result") - sens_lo.get("result"))
                       if (sens_hi.get("result") is not None and
                           sens_lo.get("result") is not None) else None, 1),
    })
    cols = active_columns(metrics, mode, psu_enable)
    # combo 列固定 [band, bw, chan] 三列；UL RB + sens Limit 已移入 active_columns，
    # 其值在 vals dict 里提供（下面补齐 ul_rb / sens_limit key）
    vals["ul_rb"] = data.get("ul_rb") or ul_rb_display_from_combo(combo_label)
    vals["sens_limit"] = data.get("sens_limit") or refsens_limit_from_combo(combo_label)
    row = [band, bw, chan] + [vals.get(c[0], "") for c in cols]
    # 判定列已置空：不再追加 PASS/FAIL 列
    for key, _label in VERDICT_COLUMNS:
        row.append("")
    return row


# ----------------------------------------------------------------------------
# 结果导出：Excel (xlsx) + TXT
# ----------------------------------------------------------------------------


def _fmt_elapsed(secs):
    """秒数 → 人可读耗时字符串：'5m39s' / '2h15m' / '45s'。"""
    try:
        s = int(round(float(secs)))
    except (TypeError, ValueError):
        return ""
    if s < 60:
        return "%ds" % s
    h, rem = divmod(s, 3600)
    m, rem2 = divmod(rem, 60)
    if h > 0:
        return "%dh%02dm%02ds" % (h, m, rem2) if rem2 else "%dh%02dm" % (h, m)
    return "%dm%02ds" % (m, rem2) if rem2 else "%dm" % m


def export_results(result, path_xlsx, path_txt):
    """把完整测量结果导出为 xlsx 与 txt。返回 (ok, err)。

    result = {"time", "cfg", "items":[{"combo","data","rows"}], "limits"}
    """
    items = result.get("items") or []
    cfg = result.get("cfg") or {}
    # 结果表按本次勾选的指标 + 灵敏度功率档动态生成列
    _metrics = cfg.get("metrics")
    _mode = cfg.get("sens_power_mode", "high")
    _psu_on = bool(cfg.get("psu_enable", True))
    exp_headers = wide_headers(_metrics, _mode, _psu_on)

    # ---- TXT ----
    try:
        lines = []
        lines.append("CMW500 LTE 射频指标测试报告（信令模式）")
        lines.append("=" * 60)
        _st = result.get("start_time") or result.get("time", "")
        _et = result.get("end_time") or result.get("time", "")
        _el = result.get("elapsed_str") or ""
        lines.append("开始时间: %s" % _st)
        lines.append("结束时间: %s" % _et)
        if _el:
            lines.append("总耗时: %s（%.1f 秒）" % (_el, result.get("elapsed_sec", 0)))
        lines.append("CMW500: %s:%s  RF口=%s  RX=%s" %
                     (cfg.get("cmw_ip", ""), cfg.get("cmw_port", ""),
                      cfg.get("rf", ""), cfg.get("rx", "")))
        lines.append("频段 / 各自带宽: %s" %
                     band_bws_summary(cfg.get("bands") or [],
                                      cfg.get("band_bws") or {}))
        lines.append("带宽并集: %s   每频段每带宽信道: 低/中/高" %
                     (",".join(bw_label(w) for w in (cfg.get("bws") or [])) or "默认"))
        lines.append("频段/带宽/信道数据源: %s" % channel_table_source())
        psu_on = cfg.get("psu_enable")
        lines.append("电源: %s  通道=%s  电压=%sV" %
                     ("启用" if psu_on else "未启用",
                      cfg.get("psu_channel", ""), cfg.get("psu_voltage", "")))
        lines.append("")
        lines.append("【测试结果（指标按列，一行一个频段/带宽/信道）】")
        lines.append(" | ".join(exp_headers))
        for it in items:
            lines.append(" | ".join(str(x) for x in it.get("wide", [])))
        lines.append("")
        # 只有勾选灵敏度时才输出探底明细 section（纯 TX 测试此 section 完全无内容）
        if _metrics and _metrics.get("sensitivity"):
            lines.append("【灵敏度探底明细】")
            lines.append("判定标准: Throughput >= %s%%（FETC:LTE:SIGN:EBLer:REL? 第2值 Total TPut）" %
                         cfg.get("sens_throughput", 95))
            for it in items:
                d = it.get("data") or {}
                for tag, sens in (("大功率", d.get("sensitivity")),
                                  ("小功率", d.get("sensitivity_low"))):
                    if sens and sens.get("table"):
                        lines.append("组合: %s [%s]" % (it["combo"], tag))
                        for rec in sens["table"]:
                            bwv = rec.get("level_bw")
                            if bwv is None:
                                bwv = epre_to_fullbw(rec["level"], "")
                            lines.append("  RS EPRE %7.1f / FullBW %8s dBm  TPut=%7s%%  %s  [%s]" %
                                         (rec["level"], fmt(bwv, 2), fmt(rec.get("throughput")),
                                          "PASS" if rec["pass"] else "FAIL",
                                          rec.get("phase", "")))
        # 报告尾部：组合数 / 总耗时
        lines.append("")
        _combo_cnt = len(items)
        _el = result.get("elapsed_str") or ""
        _el_sec = result.get("elapsed_sec", 0)
        lines.append("=" * 60)
        lines.append("完成组合: %d" % _combo_cnt)
        if _el:
            lines.append("总耗时: %s（%.1f 秒）  平均每组合: %.1f 秒" %
                         (_el, _el_sec, (_el_sec / _combo_cnt) if _combo_cnt else 0))
        with open(path_txt, "w", encoding="utf-8-sig") as f:
            f.write("\n".join(lines))
    except Exception as exc:
        return False, "TXT 写入失败: %s" % exc

    # ---- XLSX（openpyxl，缺失时降级为 CSV）----
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        wb = Workbook()
        ws = wb.active
        ws.title = "测试结果"
        ws.append(exp_headers)
        for c in ws[1]:
            c.font = Font(bold=True)

        # --- 高亮判定辅助 ---
        _FAIL_FILL = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE",
                                 fill_type="solid")  # 浅红背景
        limits_dict = {}
        # 优先从 result["cfg"] 拿用户配置的 limit（evm_lim/aclr_lim/freq_lim/sens_lim）
        try:
            limits_dict["evm_lim"] = result["cfg"].get("evm_lim", "")
            limits_dict["aclr_lim"] = result["cfg"].get("aclr_lim", "")
            limits_dict["freq_lim"] = result["cfg"].get("freq_lim", "")
        except (TypeError, KeyError):
            pass
        _headers_expanded = exp_headers  # 与 active_columns() 对齐后的实际表头

        def _cell_fill(col_idx, header, val, combo_label):
            """单元格（1-based）是否需要红色高亮。"""
            try:
                v = float(val)
            except (TypeError, ValueError):
                return None
            # EVM / 频率误差：max（≤ limit）
            if "EVM" in header or header.startswith("频率误差"):
                lim = limits_dict.get("evm_lim") if "EVM" in header else limits_dict.get("freq_lim")
                if lim:
                    try:
                        if abs(v) > float(lim):
                            return _FAIL_FILL
                    except ValueError:
                        pass
            # ACLR E-UTRA：min（≥ limit）
            elif header.startswith("ACLR E-UTRA"):
                lim = limits_dict.get("aclr_lim")
                if lim:
                    try:
                        if v < float(lim):
                            return _FAIL_FILL
                    except ValueError:
                        pass
            # 灵敏度：max（≤ 3GPP refsens limit）
            elif header.startswith("灵敏度FullBW") or header.startswith("大功率灵敏度") or header.startswith("小功率灵敏度"):
                sens_lim_str = refsens_limit_from_combo(combo_label)
                if sens_lim_str and sens_lim_str != "—":
                    try:
                        if v > float(sens_lim_str):
                            return _FAIL_FILL
                    except ValueError:
                        pass
            return None

        # 判定列已按要求去掉，结果表只写勾选指标的数值列
        for row_idx, it in enumerate(items, 2):
            wide = it.get("wide") or []
            combo_label = it.get("combo", "")
            for col_idx, header in enumerate(exp_headers, 1):
                val = wide[col_idx - 1] if col_idx - 1 < len(wide) else ""
                c = ws.cell(row=row_idx, column=col_idx, value=val)
                fill = _cell_fill(col_idx, header, val, combo_label)
                if fill is not None:
                    c.fill = fill
        # 灵敏度明细页（大功率 / 小功率两套）
        ws2 = wb.create_sheet("灵敏度探底")
        ws2.append(["组合", "功率档", "RS EPRE(dBm/15kHz)", "Full Cell BW(dBm)",
                    "Throughput(%)", "是否满足门限", "扫描阶段", "原始返回"])
        for c in ws2[1]:
            c.font = Font(bold=True)
        for it in items:
            d = it.get("data") or {}
            for tag, sens in (("大功率", d.get("sensitivity")),
                              ("小功率", d.get("sensitivity_low"))):
                if sens:
                    for rec in sens.get("table", []):
                        ws2.append([it["combo"], tag, rec["level"], rec.get("level_bw"),
                                    rec.get("throughput"), "PASS" if rec["pass"] else "FAIL",
                                    rec.get("phase", ""), rec["raw"]])
        # 配置页
        ws3 = wb.create_sheet("测试配置")
        ws3.append(["项目", "值"])
        for c in ws3[1]:
            c.font = Font(bold=True)
        _st = result.get("start_time") or result.get("time", "")
        _et = result.get("end_time") or result.get("time", "")
        _el = result.get("elapsed_str") or ""
        ws3.append(["开始时间", _st])
        ws3.append(["结束时间", _et])
        if _el:
            ws3.append(["总耗时", "%s (%.1f 秒)" % (_el, result.get("elapsed_sec", 0))])
        for k in ("cmw_ip", "cmw_port", "rf", "rx", "enp", "eatt", "rsrp",
                  "pmax", "freq", "psu_enable", "psu_channel",
                  "psu_voltage", "sens_start", "sens_stop", "sens_step",
                  "sens_near", "sens_coarse", "sens_throughput", "stats"):
            ws3.append([k, cfg.get(k, "")])
        ws3.append(["bands", ",".join(band_label(b) for b in (cfg.get("bands") or []))])
        ws3.append(["bws", ",".join(bw_label(w) for w in (cfg.get("bws") or []))])
        _bb = cfg.get("band_bws") or {}
        for _b in (cfg.get("bands") or []):
            ws3.append(["band %s" % band_label(_b),
                        ",".join(bw_label(_w) for _w in (_bb.get(_b) or []))])
        ws3.append(["channel_source", channel_table_source()])
        wb.save(path_xlsx)
    except ImportError:
        # 降级：用 CSV 内容写入 .csv（同名），并提示
        import csv as _csv
        alt = os.path.splitext(path_xlsx)[0] + ".csv"
        with open(alt, "w", newline="", encoding="utf-8-sig") as f:
            w = _csv.writer(f)
            w.writerow(exp_headers)
            for it in items:
                w.writerow(it.get("wide") or [])
        return True, "openpyxl 未安装，已改为 CSV: %s" % alt
    except Exception as exc:
        return False, "XLSX 写入失败: %s" % exc

    return True, ""


# ----------------------------------------------------------------------------
# 自检 / 演示（无仪表）
# ----------------------------------------------------------------------------


def demo_measurement(cfg):
    """生成演示测量结果（与真机结构一致），用于无仪表验证界面与导出。"""
    # 首位 0 = Reliability Indicator OK（手册 p.1736，与真机 FETCh 结构一致）
    ma = [0, 0, 2.31, 2.55, 4.10, 4.52, 1.92, 2.14, 3.01, 3.28,
          1.21, 1.33, 2.02, 2.19, -32.5, -12.3, 0.42, 22.80, 23.10,
          -61.30, 1.80, 2.02, 1.70, 1.83, 0.91, 1.02]
    mx = [0, 0, 3.90, 4.21, 5.62, 5.91, 3.11, 3.44, 4.52, 4.78,
          2.10, 2.31, 3.12, 3.40, -29.8, -25.4, 0.78, 22.6, 23.2,
          7.80, 8.42, -62.10, -60.30, 2.91, 3.12, 1.61, 1.80, 1.55, 1.74]
    ac = [0, 42.1, 45.3, 42.5, 8.02, 43.1, 45.0, 42.8]
    data = {
        "state": "RDY", "mod_avg": ma, "mod_ext": mx, "aclr": ac,
        "raw": {"mod_avg": ",".join("%.6g" % v for v in ma),
                "aclr": ",".join("%.6g" % v for v in ac)},
    }
    if cfg.get("psu_enable", True):
        # 源头已存 mA，demo 数据同步
        data["psu"] = {"channel": cfg.get("psu_channel", "P6V"),
                       "current": 482.0, "current_min": 470.0, "current_max": 495.0,
                       "idle_current": 31.0, "voltage": 3.800, "power": 1.832,
                       "samples": 3}
    if cfg.get("metrics", {}).get("sensitivity", True):
        data["sensitivity"] = {
            # RS EPRE -124.8 @10MHz -> Full Cell BW ≈ -97.0 dBm
            "result": round(-124.8 + full_cell_bw_offset_db("B100"), 2),
            "result_epre": -124.8,
            "threshold": float(cfg.get("sens_throughput", 95)),
            "criterion": "Throughput >= %s%%" % cfg.get("sens_throughput", 95),
            "level_label": "Full Cell BW Pow (dBm)",
            "table": [
                {"level": -117.8, "level_bw": round(-117.8 + full_cell_bw_offset_db("B100"), 2), "throughput": 100.0, "state": "RDY", "pass": True, "phase": "coarse", "raw": "0,100,0,0,100,..."},
                {"level": -122.8, "level_bw": round(-122.8 + full_cell_bw_offset_db("B100"), 2), "throughput": 97.5, "state": "RDY", "pass": True, "phase": "fine", "raw": "0,97,3,0,97.5,..."},
                {"level": -124.8, "level_bw": round(-124.8 + full_cell_bw_offset_db("B100"), 2), "throughput": 96.0, "state": "RDY", "pass": True, "phase": "fine", "raw": "0,96,4,0,96,..."},
                {"level": -125.8, "level_bw": round(-125.8 + full_cell_bw_offset_db("B100"), 2), "throughput": 86.0, "state": "RDY", "pass": False, "phase": "fine", "raw": "0,86,14,0,86,..."},
            ],
        }
    return data


def run_selftest():
    # 解析
    assert parse_values("1,2.5,-3.1E+1,NCAP,0.5") == [1.0, 2.5, -31.0, None, 0.5]
    assert parse_scalar("NCAP,3.8") == 3.8
    # INV/NAV 无效值解析（UE 未注册时 FETCH 返回 4,INV,INV,...）
    assert parse_values("4,INV,INV,NCAP") == [4.0, None, None, None]
    # Reliability 指示码（手册 p.1736）：首位单独解析用于精确诊断
    assert parse_reliability("4,INV,INV,INV")[0] == 4
    assert "Underdriven" in parse_reliability("4,INV,INV")[1]
    assert parse_reliability("0,0,10.4,-17.5")[0] == 0
    assert "OK" in parse_reliability("0,0,10.4")[1]
    assert parse_reliability("6,INV")[0] == 6
    assert parse_reliability("bogus,1")[0] is None
    # 诊断文案：Underdriven(4) 要给出上行信号/期望功率/PUSCH 等针对性排查
    diag4 = diagnose_reliability(4, {"enp": "23", "rsrp": "-80"})
    assert "ENPower" in diag4 and "PUSCH" in diag4 and "23" in diag4
    assert "CSPath" in diagnose_reliability(6)
    # 连接状态判定：ON/SIGN/OFF 不算注册；ATTached/ATT、CESTablished/CEST 都算
    assert not conn_ready("ON")
    assert not conn_ready("OFF")
    assert not conn_ready("SIGN")
    assert not conn_ready("CONNecting, signaling in progress")
    assert conn_ready("ATTached")
    assert conn_ready("ATT")          # 仪表固件实际返回的缩写
    assert conn_ready("CESTablished")
    assert conn_ready("CEST")
    # 测量前注册复查：已注册立即通过；未注册且等待超时 -> 抛 ScpiError 中止
    class _RegScpi:
        def __init__(self, state):
            self.state = state
        def query(self, cmd, timeout=None, expect=None):
            return self.state
    assert ensure_registered(_RegScpi("ATT")) == "ATT"
    try:
        ensure_registered(_RegScpi("ON"), timeout=0.15, interval=0.05)
        assert False, "未注册应抛 ScpiError 中止"
    except ScpiError:
        pass
    # 小区判定：ON,PEND 未就绪；ON,ADJ 就绪
    assert not cell_ready("ON,PEND")
    assert cell_ready("ON,ADJ")
    # 判定列已去掉：宽表/表头不再含任何"判定"列
    assert len(VERDICT_COLUMNS) == 0
    assert all("判定" not in h for h in wide_headers())
    # 信令初始化命令
    cfg = dict(CONFIG_DEFAULTS)
    # 自检用 0 稳定时间，避免 4s×N 次假仪表测量拖慢自检（真机仍用 4s 默认值）
    cfg.update({"band": "OB2", "eatt": 2.0, "tpc_settle": "0"})
    cmds, meta = build_signaling_init(cfg, band="OB2", bw="B100")
    joined = "\n".join(cmds)
    for token in ("*RST", "ROUT:LTE:SIGN:SCEN:SCEL RF1C,RX1,RF1C,TX1",
                  "CONF:LTE:SIGN:BAND OB2", "CONF:LTE:SIGN:CELL:BAND:DL B100",
                  "CONF:LTE:SIGN:UL:PMAX 23.0", "CONF:LTE:SIGN:CONN:UET RMC",
                  "CONF:LTE:SIGN:CONN:STYP RMC", "CONF:LTE:SIGN:RFS:ENPMode ULPC",
                  "ROUT:LTE:MEAS:SCEN:CSPath 'LTE Sig1'"):
        assert token in joined, token
    # 外部衰减成对补偿（手册 p.1729：OUTP 与 INP 成对，FDD 双工共用衰减器）
    assert "CONF:LTE:SIGN:RFS:EATT:OUTP 2.0" in joined, joined
    assert "CONF:LTE:SIGN:RFS:EATT:INP 2.0" in joined, joined
    # FDC 频率响应校准（ASRrun V2.0.7 对比）——11 条命令：RCL + DELete + CREATE + 5×ADD + ACTivate TX + ACTivate RX
    assert "CONF:BASE:FDC:RCL" in joined, "FDC RCL"
    assert "CONF:BASE:FDC:CTABle:CREATE 'losstable', 450000000, 0.3" in joined, "FDC CREATE"
    assert "CONF:BASE:FDC:CTABle:ADD 'losstable', 1000000000, 0.5" in joined, "FDC ADD"
    assert "CONFigure:FDCorrection:ACTivate RF1C, 'losstable', TX" in joined, "FDC ACT TX"
    assert "CONFigure:FDCorrection:ACTivate RF1C, 'losstable', RX" in joined, "FDC ACT RX"
    # RBP 固定 LOW（ASRrun 参考流程），避免切带宽/STYP 后位置残留
    assert "CONF:LTE:SIGN:CONN:RMC:RBP:DL LOW" in joined, "RBP DL"
    assert "CONF:LTE:SIGN:CONN:RMC:RBP:UL LOW" in joined, "RBP UL"
    # 首组合带宽决定初始化满 RB：1.4MHz→N6（不能沿用 cfg 默认 N50）
    _c14, _ = build_signaling_init(dict(cfg), band="OB5", bw="B014")
    assert "CONF:LTE:SIGN:CONN:RMC:DL N6,QPSK,KEEP" in "\n".join(_c14), "RB B014"
    assert "CONF:LTE:SIGN:CONN:RMC:UL N6,QPSK,KEEP" in "\n".join(_c14), "RB B014 UL"
    _c30, _ = build_signaling_init(dict(cfg), band="OB5", bw="B030")
    assert "CONF:LTE:SIGN:CONN:RMC:DL N15,QPSK,KEEP" in "\n".join(_c30), "RB B030"
    # ---- TDD 支持 ----
    assert band_duplex("OB2") == "FDD" and band_duplex("OB5") == "FDD"
    assert band_duplex("OB34") == "TDD" and band_duplex("OB38") == "TDD"
    assert band_duplex("OB41") == "TDD" and is_tdd_band("OB39")
    assert not tdd_cell_commands(cfg, "OB2"), "FDD 不应写 TDD 帧结构"
    _tdd_cmds = tdd_cell_commands(dict(cfg, tdd_uldl="2", tdd_ssub="1"), "OB38")
    assert _tdd_cmds == ["CONF:LTE:SIGN:DMODe TDD",
                         "CONF:LTE:SIGN:CELL:ULDL 2",
                         "CONF:LTE:SIGN:CELL:SSUBframe 1"], _tdd_cmds
    assert fdd_cell_commands() == ["CONF:LTE:SIGN:DMODe FDD"]
    # 边界钳制：ULDL 0~6、SSUB 0~9
    assert tdd_cell_commands(dict(cfg, tdd_uldl="99", tdd_ssub="-5"), "OB38") == [
        "CONF:LTE:SIGN:DMODe TDD", "CONF:LTE:SIGN:CELL:ULDL 6",
        "CONF:LTE:SIGN:CELL:SSUBframe 0"]
    # FDD 首建显式 DMODe FDD；TDD 首建 DMODe TDD + 帧结构，且 BAND 在其后
    _cfdd, _ = build_signaling_init(dict(cfg), band="OB2", bw="B100")
    assert "CONF:LTE:SIGN:DMODe FDD" in _cfdd
    assert _cfdd.index("CONF:LTE:SIGN:DMODe FDD") < \
        _cfdd.index("CONF:LTE:SIGN:BAND OB2")
    _ctdd, _ = build_signaling_init(dict(cfg), band="OB38", bw="B100")
    assert "CONF:LTE:SIGN:DMODe TDD" in _ctdd
    assert "CONF:LTE:SIGN:CELL:ULDL 1" in _ctdd
    assert "CONF:LTE:SIGN:CELL:SSUBframe 7" in _ctdd
    assert _ctdd.index("CONF:LTE:SIGN:DMODe TDD") < \
        _ctdd.index("CONF:LTE:SIGN:BAND OB38")
    # retune_cell 重建：duplex 命令必须在 CELL OFF 之后、BAND 之前
    class _TddScpi:
        def __init__(self):
            self.writes = []
        def write(self, cmd):
            self.writes.append(cmd)
        def query(self, cmd, timeout=None, expect=None):
            if "CELL:STAT:ALL" in cmd: return "ON,ADJ"
            if "PSW:STAT" in cmd: return "ATT"
            return ""
    _td = _TddScpi()
    retune_cell(_td, dict(cfg), "OB38", "B100", 37800)
    w = _td.writes
    assert w.index("SOUR:LTE:SIGN:CELL:STAT OFF") < \
        w.index("CONF:LTE:SIGN:DMODe TDD") < \
        w.index("CONF:LTE:SIGN:BAND OB38"), w[:12]
    # 授权探测：mock -220 → 判不支持；无错误 → 支持且恢复 FDD
    class _CapScpi:
        def __init__(self, err, cell="OFF"):
            self.err = err; self.cell = cell; self.writes = []; self.q = 0
        def write(self, cmd):
            self.writes.append(cmd)
            if "CELL:STAT ON" in cmd: self.cell = "ON"
            if "CELL:STAT OFF" in cmd: self.cell = "OFF"
        def query(self, cmd, timeout=None, expect=None):
            if "DMODe?" in cmd: return "FDD"
            if "CELL:STAT?" in cmd: return self.cell
            if "SYST:ERR?" in cmd:
                self.q += 1
                return self.err if self.q == 1 else '0,"No error"'
            return ""
    _no = _CapScpi('-220,"Parameter error;access permission"')
    ok, detail = detect_tdd_capability(_no)
    assert not ok and "-220" in detail, (ok, detail)
    assert "CONF:LTE:SIGN:DMODe FDD" in _no.writes
    _yes = _CapScpi('0,"No error"')
    ok, detail = detect_tdd_capability(_yes)
    assert ok and "可设置" in detail, (ok, detail)
    assert "CONF:LTE:SIGN:DMODe FDD" in _yes.writes  # 恢复原 duplex
    # CELL ON 时探测必须先关小区、结束恢复 ON
    _on = _CapScpi('0,"No error"', cell="ON")
    assert detect_tdd_capability(_on)[0] is True
    assert "SOUR:LTE:SIGN:CELL:STAT OFF" in _on.writes
    assert "SOUR:LTE:SIGN:CELL:STAT ON" in _on.writes
    # 显式 MEV 配置：只开 TXM + 按需 ACLR，关 SEMask/IEMissions
    _mev = build_mev_config_commands(dict(cfg, metrics={"maxpower": True, "aclr": True}))
    assert "CONF:LTE:MEAS:MEV:REP SING" in _mev
    assert "CONF:LTE:MEAS:MEV:RES:TXM ON" in _mev
    assert "CONF:LTE:MEAS:MEV:RES:ACLR ON" in _mev
    assert "CONF:LTE:MEAS:MEV:RES:SEMask OFF" in _mev
    # MSUBframes 偏移按双工区分（参考工具 1756.log）：FDD=0，TDD=2
    assert "CONF:LTE:MEAS:MEV:MSUBframes 0,10,0" in _mev
    _mev_tdd = build_mev_config_commands(
        dict(cfg, metrics={"maxpower": True, "aclr": True}), "OB38")
    assert "CONF:LTE:MEAS:MEV:MSUBframes 0,10,2" in _mev_tdd
    assert "CONF:LTE:MEAS:MEV:MSUBframes 0,10,0" not in _mev_tdd
    assert meta["cell_on"].startswith("SOUR:LTE:SIGN:CELL:STAT ON")
    assert meta["conn_cmd"].startswith("CALL:LTE:SIGN:PSW:ACT CONN")
    # 满功率命令（PMAX 只是上限，必须 TPC MAXPower 才推到满功率）
    mp = build_force_maxpower_cmd(dict(cfg, tpc_settle="2.5"))
    assert mp["set_max"] == "CONF:LTE:SIGN:UL:TPC:SET MAXPower", mp["set_max"]
    assert mp["settle"] == 2.5, mp["settle"]
    # 默认稳定时间 4s（缓解 Rel=3 Overdriven），配置缺失时也回退 4s
    assert CONFIG_DEFAULTS["tpc_settle"] == "4"
    assert build_force_maxpower_cmd({})["settle"] == 4.0
    assert CONFIG_DEFAULTS["sens_min_total"] == "50"
    # 功率档独立 Total 门限 + 切带宽后上行稳定等待
    assert CONFIG_DEFAULTS["sens_min_total_low"] == "20"
    assert sens_min_total_for({}) == 50.0
    assert sens_min_total_for({}, "low") == 20.0
    assert sens_min_total_for({"sens_min_total_low": "35"}, "low") == 35.0
    assert CONFIG_DEFAULTS["attach_settle"] == "2"
    # 方案 A：掉线特征门限 + 重附着快速升级窗口
    assert CONFIG_DEFAULTS["sens_drop_total"] == "10"
    assert CONFIG_DEFAULTS["attach_quick_timeout"] == "5"
    assert CONFIG_DEFAULTS["fast_retune_timeout"] == "8"
    assert CONFIG_DEFAULTS["fast_retune_passive_wait"] == "30"  # 热切掉线无AT口被动等
    assert CONFIG_DEFAULTS["tdd_uldl"] == "1" and CONFIG_DEFAULTS["tdd_ssub"] == "7"
    assert CONFIG_DEFAULTS["tdd_detect"] == "0"
    assert CONFIG_DEFAULTS["fast_retune_busy_wait"] == "25"     # 首窗口未确认的忙期宽限确认
    assert CONFIG_DEFAULTS["init_popup_ms"] == "3000"
    assert CONFIG_DEFAULTS["at_port"] == ""   # AT 口默认 auto，启动自动识别
    assert CONFIG_DEFAULTS["at_cmd"] == ""    # AT 指令输入框默认空，关闭/发送时持久化
    assert CONFIG_DEFAULTS["at_auto_recover"] == "1"   # 掉线 AT CFUN 自动恢复默认启用
    assert CONFIG_DEFAULTS["at_recover_wait"] == "60"  # CFUN=1 后回网等待（SIGN 抖动 3 轮提前判负）
    # _at_recovery_port 解析优先级：显式 at_port > 探测口 at_port_detected > None
    assert _at_recovery_port({"at_port": "COM4"}) == "COM4"
    assert _at_recovery_port({"at_port": "auto", "at_port_detected": "COM4"}) == "COM4"
    assert _at_recovery_port({"at_port": "auto"}) is None
    assert _at_recovery_port({"at_port": "COM4", "at_auto_recover": "0"}) is None
    assert _at_recovery_port({}) is None
    # 盲扫命令
    b = build_bler_commands(cfg, -80.0)
    assert b["set_level"] == "CONF:LTE:SIGN:DL:RSEPre:LEVel -80.0"  # set_level 是扫描时动态传入的
    assert b["subframes"] == "CONF:LTE:SIGN:EBLer:SFRames 600"  # ASRrun V2.0.7 对比优化：1000→600
    assert b["fetch"] == "FETC:LTE:SIGN:EBLer:REL?"  # ASRrun V2.0.7 用 REL 格式直接返回 %BLER
    # UL RB 配置（3GPP 表 7.3.5-2）
    assert ul_rb_for("OB1", "B100") == (50, 0)
    assert ul_rb_for("OB2", "B150") == (50, 25)
    assert ul_rb_for("OB20", "B200") == (20, 16)   # 中段起始 RB，RMC 表示不了，UDCH 精确
    ud = build_udch_commands("OB2", "B150")
    assert ud["ul"] == "CONF:LTE:SIGN:CONN:UDCH:UL 50,25,QPSK,KEEP", ud["ul"]
    assert ud["dl"] == "CONF:LTE:SIGN:CONN:UDCH:DL 75,0,QPSK,KEEP", ud["dl"]
    assert ud["set_udch"] == "CONF:LTE:SIGN:CONN:STYP UDCH"
    assert restore_rmc_commands(cfg)[0] == "CONF:LTE:SIGN:CONN:STYP RMC"
    # 灵敏度测完按当前 bw 恢复满 RB（不能沿用全局 N50）
    assert "CONF:LTE:SIGN:CONN:RMC:UL N6,QPSK,KEEP" in restore_rmc_commands(cfg, "B014")
    assert "CONF:LTE:SIGN:CONN:RMC:UL N15,QPSK,KEEP" in restore_rmc_commands(cfg, "B030")
    # 恢复 RMC 时必须关 OCNG，否则后续 TX 会 Overdriven(Rel=3)
    assert "CONF:LTE:SIGN:DL:OCNG OFF" in restore_rmc_commands(cfg, "B014")
    # RX 灵敏度在 RMC 下配置（3GPP §7.3.5 + ASRrun 参考）；TX full RB 独立由恢复命令负责
    class _SensRbScpi:
        def __init__(self, err=""):
            self.w = []
            self.err = err
        def write(self, cmd):
            self.w.append(cmd)
        def query(self, cmd, timeout=None, expect=None):
            return self.err
    # 表 7.3.5-2：B5 1.4/3/5MHz RBstart=0 → RMC；10MHz 25@25 → UDCH
    _rb14 = _SensRbScpi()
    _b14 = apply_sens_rb(_rb14, "OB5", "B014", cfg)
    assert _b14["mode"] == "RMC" and _b14["rb_applied"] == (6, 0), _b14
    assert "CONF:LTE:SIGN:CONN:UET RMC" in _rb14.w
    assert "CONF:LTE:SIGN:CONN:STYP RMC" in _rb14.w
    assert "CONF:LTE:SIGN:CONN:RMC:UL N6,QPSK,KEEP" in _rb14.w
    assert "CONF:LTE:SIGN:CONN:RMC:RBP:UL LOW" in _rb14.w
    assert "CONF:LTE:SIGN:CONN:STYP UDCH" not in _rb14.w
    assert "CONF:LTE:SIGN:CONN:UDCH:UL" not in _rb14.w
    _rb50 = _SensRbScpi()
    _b50 = apply_sens_rb(_rb50, "OB5", "B050", cfg)
    assert _b50["mode"] == "RMC" and _b50["rb_applied"] == (25, 0), _b50
    assert "CONF:LTE:SIGN:CONN:RMC:UL N25,QPSK,KEEP" in _rb50.w
    # 10MHz 25@25：RBstart>0 必须 UDCH，不得再发 RMC RBP P25
    _rb10 = _SensRbScpi()
    _b10 = apply_sens_rb(_rb10, "OB5", "B100", cfg)
    assert _b10["mode"] == "UDCH" and _b10["rb_applied"] == (25, 25), _b10
    assert "CONF:LTE:SIGN:CONN:STYP UDCH" in _rb10.w
    assert "CONF:LTE:SIGN:CONN:UDCH:UL 25,25,QPSK,KEEP" in _rb10.w
    assert "CONF:LTE:SIGN:CONN:RMC:RBP:UL P25" not in _rb10.w
    # start=0 的其它频段（B74 10MHz 25@0）仍在 RMC
    _rb74 = _SensRbScpi()
    _b74 = apply_sens_rb(_rb74, "OB74", "B100", cfg)
    assert _b74["mode"] == "RMC" and _b74["rb_applied"] == (25, 0), _b74
    # OCNG ON：灵敏度段使 Full Cell BW Power 恒定（ASRrun 参考流程）
    _mev_sens = build_bler_commands(cfg, -85.0)
    assert _mev_sens["fetch"] == "FETC:LTE:SIGN:EBLer:REL?"
    # 灵敏度三级扫描：0.1dBm 精度 + 起始 -85 + 中步
    assert CONFIG_DEFAULTS["sens_start"] == "-85"
    assert CONFIG_DEFAULTS["sens_step"] == "0.1"
    assert CONFIG_DEFAULTS["sens_near"] == "1"
    # 灵敏度判定：Throughput >= 95%，判定量 FETC:EBLer:REL? 第5值(idx 4)
    assert CONFIG_DEFAULTS["sens_throughput"] == "95"
    assert "sens_bler" not in CONFIG_DEFAULTS
    # 热点窗口收窄（A+B 加速）：margin/span 默认 6/4 -> 3/3
    assert CONFIG_DEFAULTS["sens_hotspot_margin"] == "3"
    assert CONFIG_DEFAULTS["sens_hotspot_span"] == "3"
    # 3GPP 理论 REFSENS 种子：B1/10MHz 理论 -96.3dBm FullBW -> RS EPRE ≈ -124.08
    assert abs(refsens_epre_for("OB1", "B100") - (-96.3 - 27.78)) < 0.02
    assert refsens_epre_for("OB1", "B014") is None  # B1 不支持 1.4MHz
    assert refsens_epre_for("OB99", "B100") is None  # 无数据
    # 3GPP 灵敏度 Limit 结果列：FullBW 口径、从组合显示名反查、缺失回退
    assert refsens_limit_for("OB1", "B100") == -96.3
    assert refsens_limit_for("OB5", "B014") == -102.5
    assert refsens_limit_display("OB1", "B100") == "-96.3"
    assert refsens_limit_from_combo("B1 / 10MHz / 低(CH50)") == "-96.3"
    assert refsens_limit_from_combo("B5 / 1.4MHz / 低(CH2407)") == "-102.5"
    assert refsens_limit_from_combo("B99 / 10MHz / 低(CH50)") == "—"
    # reconnect_after_styp 主动重连：已注册零开销；掉线发 ACT CONN
    class _StypScpi:
        """初始状态 init_st；仅当收到 ACT CONN 后才转 act_st（模拟真实时序）。"""
        def __init__(self, init_st, act_st="ATT"):
            self.st = init_st
            self.act_st = act_st
            self.writes = []
        def query(self, cmd, timeout=None, expect=None):
            return self.st
        def write(self, cmd):
            self.writes.append(cmd)
            if "PSW:ACT CONN" in cmd:
                self.st = self.act_st
    _s = _StypScpi("ATT")
    assert reconnect_after_styp(_s) == "ATT" and not _s.writes  # 已注册：不碰电平、零写入
    _s = _StypScpi("ON")
    assert reconnect_after_styp(_s, cfg={"rsrp": "-80"}) == "ATT"
    assert "CALL:LTE:SIGN:PSW:ACT CONN" in _s.writes
    # 掉线重附着：必须先拉回强信号（cfg.rsrp），且电平命令先于 ACT CONN
    _lvl_cmd = "CONF:LTE:SIGN:DL:RSEPre:LEVel -80.0"
    assert _lvl_cmd in _s.writes, _s.writes
    assert _s.writes.index(_lvl_cmd) < _s.writes.index("CALL:LTE:SIGN:PSW:ACT CONN")
    # 未传 cfg 时回退默认 -85
    _s2 = _StypScpi("ON")
    reconnect_after_styp(_s2)
    assert "CONF:LTE:SIGN:DL:RSEPre:LEVel -85.0" in _s2.writes
    # 方案 A：硬掉线（ACT CONN 快速窗口内持续 ON）→ 自动重建小区，不再死等 timeout×2
    class _HardDropScpi:
        def __init__(self):
            self.states = ["ON"]
            self.writes = []
            self.cell_rebuilt = False
        def query(self, cmd, timeout=None, expect=None):
            if "CELL:STAT:ALL" in cmd:
                return "ON,ADJ"
            if "PSW:STAT" in cmd:
                if self.cell_rebuilt:
                    return "ATT"
                return self.states.pop(0) if self.states else "ON"
            return "ON"
        def write(self, cmd):
            self.writes.append(cmd)
            if cmd == "SOUR:LTE:SIGN:CELL:STAT ON":
                self.cell_rebuilt = True
    _hd = _HardDropScpi()
    _fast_cfg = dict(CONFIG_DEFAULTS, attach_quick_timeout="1", attach_settle="0")
    _hd_st = reconnect_after_styp(_hd, cfg=_fast_cfg,
                                  retune_ctx={"band": "OB5", "bw": "B014", "earfcn": 25250})
    assert _hd_st == "ATT", _hd_st
    assert "CALL:LTE:SIGN:PSW:ACT CONN" in _hd.writes
    assert _hd.writes.count("SOUR:LTE:SIGN:CELL:STAT OFF") == 1, _hd.writes
    assert "SOUR:LTE:SIGN:CELL:STAT ON" in _hd.writes
    # 频段/带宽快切自检（连接态热切模式，参考工具 18340.log 实证）：
    # 同频段热切带宽 → UE 保持注册即 True；不写 BAND、不重发 CELL ON、不关小区；
    # UE 掉线（PSW=ON）→ 有 AT 口立即 False（交 AT 恢复）/无 AT 口被动等后 False。
    class _FastScpi:
        def __init__(self, st="ATT"):
            self.st = st
            self.writes = []
        def write(self, cmd):
            self.writes.append(cmd)
        def query(self, cmd, timeout=None, expect=None):
            if "PSW:STAT" in cmd:
                return self.st
            return ""
    _fcfg = dict(CONFIG_DEFAULTS, fast_retune_timeout="1",
                 fast_retune_passive_wait="1", fast_retune_busy_wait="1")
    # 热切成功（同频段换带宽，UE 保持 ATT）：不重启小区
    _f_on = _FastScpi("ATT")
    assert try_fast_retune(_f_on, _fcfg, "OB5", "B100", 25250,
                           last_cell=("OB5", "B014")) is True
    assert "CONF:LTE:SIGN:BAND OB5" not in _f_on.writes        # 同频段跳过 BAND 写
    assert "SOUR:LTE:SIGN:CELL:STAT ON" not in _f_on.writes    # 关键：不重启小区
    assert "CONF:LTE:SIGN:CELL:BAND:DL B100" in _f_on.writes
    assert "CONF:LTE:SIGN:CONN:RMC:DL N50,QPSK,KEEP" in _f_on.writes
    assert "SOUR:LTE:SIGN:CELL:STAT OFF" not in _f_on.writes   # 关键：不关小区
    assert "CALL:LTE:SIGN:PSW:ACT CONN" not in _f_on.writes    # 关键：不主动附着
    # 带宽增大方向两段式（2026-11-25 用户方案）：先停靠中间信道 CH2525（跨带宽
    # 恒为频段中心，旧/新带宽下均合法），UE 确认回网后再做同中心带宽切换
    # （参考工具实证不掉线的模式）→ 新 RMC → 目标信道
    assert _f_on.writes[0] == "CONF:LTE:SIGN:RFS:CHAN:DL 2525", _f_on.writes
    _i_bw = _f_on.writes.index("CONF:LTE:SIGN:CELL:BAND:DL B100")
    _i_rmc = next(i for i, c in enumerate(_f_on.writes) if "RMC:DL N50" in c)
    _i_tchan = _f_on.writes.index("CONF:LTE:SIGN:RFS:CHAN:DL 25250")
    assert 0 < _i_bw < _i_rmc < _i_tchan, _f_on.writes
    # 换频段（无 last_cell）：仍写 BAND + 重发 CELL ON（RF 重调场景）
    _f_band = _FastScpi("ATT")
    assert try_fast_retune(_f_band, _fcfg, "OB5", "B100", 25250) is True
    assert "CONF:LTE:SIGN:BAND OB5" in _f_band.writes
    assert "SOUR:LTE:SIGN:CELL:STAT ON" in _f_band.writes
    # 热切后 UE 掉线（PSW=ON）→ 无 AT 口被动等 1s 仍 ON → False（回退重建）
    # 10M→1.4M 减小方向：写入顺序 RMC → 带宽 → 信道（先降 RB 避免中间态超限）
    _f_off = _FastScpi("ON")
    assert try_fast_retune(_f_off, _fcfg, "OB5", "B014", 25250,
                           last_cell=("OB5", "B100")) is False
    assert "CONF:LTE:SIGN:CONN:RMC:UL N6,QPSK,KEEP" in _f_off.writes  # 1.4M RB 重配
    _i_bw2 = _f_off.writes.index("CONF:LTE:SIGN:CELL:BAND:DL B014")
    _i_rmc2 = next(i for i, c in enumerate(_f_off.writes) if "RMC:UL N6" in c)
    _i_chan2 = next(i for i, c in enumerate(_f_off.writes) if "CHAN:DL" in c)
    assert _i_rmc2 < _i_bw2 < _i_chan2, _f_off.writes
    # 热切后 UE 掉线 + 有 AT 口 → 立即 False 交调用方 AT 恢复（不做被动等待）
    _f_at = _FastScpi("ON")
    _fcfg_at = dict(_fcfg, at_port="COM4")
    assert try_fast_retune(_f_at, _fcfg_at, "OB5", "B050", 25250,
                           last_cell=("OB5", "B014")) is False
    # 无 AT 口：UE 短暂掉线后自主重附着（PSW 前 2 次 ON，第 3 次起 ATT）
    # → 停靠/首窗口未确认、宽限确认期内拿到 ATT → True（参考工具实测自恢复 ~26s）
    class _FastBusy(_FastScpi):
        def __init__(self):
            _FastScpi.__init__(self, "ON")
            self.psw_polls = 0
        def query(self, cmd, timeout=None, expect=None):
            if "PSW:STAT" in cmd:
                self.psw_polls += 1
                return "ATT" if self.psw_polls >= 3 else "ON"
            return ""
    _f_busy = _FastBusy()
    assert try_fast_retune(_f_busy, _fcfg, "OB5", "B100", 25250,
                           last_cell=("OB5", "B014")) is True, "UE 自主回网应成功"
    assert "SOUR:LTE:SIGN:CELL:STAT OFF" not in _f_busy.writes  # 未回退重建
    # 理论种子扫描（假仪表）：Throughput 在 knee 处由 100% 跌到 0%；理论种子
    # -124.08，应从 seed+margin(-121.08) 起步（而非 -85），二分精扫探到 knee。
    class _SensScpi:
        """模拟 CMW500 FETCh:EBLer:REL? 单点查询（6 值，2026-11-25 实测格式）。

        响应: Rel(0), TotalTPut%(1), DTX%(2), BLER%(3), 副统计(4), 副统计(5)
        idx3 = 100 - idx1。tp_pass/tp_fail 可设边界值，验证「Throughput >= 门限」
        语义：恰为 95.0% 时应判 PASS，94.9% 应判 FAIL。
        invalid_below：电平低于该值时返回 "0,0,,,,"（无统计块，BLER 缺失）；
        None = 永远有统计。flaky_q：第 N 次 REL? 查询返回无统计块（重测恢复场景）。
        """
        def __init__(self, knee, tp_pass=100.0, tp_fail=0.0, invalid_below=None,
                     psw_on_at=None, flaky_q=None):
            self.knee = knee
            self.levels = []
            self.inits = 0
            self.tp_pass = tp_pass
            self.tp_fail = tp_fail
            self.invalid_below = invalid_below
            self.flaky_q = flaky_q
            # 第 N 次 PSW:STAT? 查询返回 ON（模拟扫描中掉线）；None=始终 ATT。
            # 初始 recheck 和末尾恢复都返回 ATT，避免误触发重附着流程。
            self.psw_on_at = psw_on_at
            self.psw_q = 0
            self._rel_q = 0
        def write(self, cmd):
            if "RSEPre:LEVel" in cmd:
                self.levels.append(float(cmd.rsplit(" ", 1)[1]))
            if "INIT:LTE:SIGN:EBLer" in cmd:
                self.inits += 1
        def query(self, cmd, timeout=None, expect=None):
            if "EBLer:STAT?" in cmd:
                return "RDY"
            if "EBLer:REL?" in cmd:
                self._rel_q += 1
                cur = self.levels[-1] if self.levels else -85.0
                if self.flaky_q is not None and self._rel_q == self.flaky_q:
                    return "0,0,,,,"   # 无统计块（BLER 缺失 → TPut 不可信）
                if self.invalid_below is not None and cur < self.invalid_below:
                    return "0,0,,,,"   # 无统计块
                tput = self.tp_pass if cur >= self.knee else self.tp_fail
                ber = 100.0 - tput
                # idx1=TotalTPut, idx2=DTX, idx3=BLER, idx4/5 副统计
                return "0,%s,0.000000E+000,%s,0.000000E+000,0.000000E+000" % (
                    "%E" % tput, "%E" % ber)
            if "PSW:STAT?" in cmd:
                self.psw_q += 1
                return "ON" if self.psw_q == self.psw_on_at else "ATT"
            if "SYST:ERR?" in cmd:
                return '0,"No error"'
            return ""
    seed12 = refsens_epre_for("OB1", "B100")   # ≈ -124.08
    sens_scpi = _SensScpi(knee=-124.0)
    sres = measure_sensitivity(sens_scpi, dict(cfg, sensitivity="x"),
                               {"sensitivity"}, band="OB1", bw="B100",
                               theory_seed=seed12)
    # 首个扫描电平应为 seed+seed_margin（-121.08）；注意 warmup 阶段先设了 -85
    _scan_levels = [lv for lv in sens_scpi.levels if lv != -85.0]
    assert _scan_levels and abs(_scan_levels[0] - (seed12 + float(cfg.get("sens_seed_margin") or cfg["sens_hotspot_margin"]))) < 0.05, \
        (_scan_levels[:3], sens_scpi.levels[:5])
    assert sres and sres["result_epre"] is not None
    assert abs(sres["result_epre"] - (-124.0)) <= 0.11, sres["result_epre"]
    # 判定标准 = Throughput >= 95%；明细记录 TPut 且与 pass 一致
    assert sres["criterion"] == "Throughput >= 95.0%", sres.get("criterion")
    assert sres["threshold"] == 95.0
    for _rec in sres["table"]:
        assert "raw_rel" not in _rec and "total" not in _rec, _rec
        assert _rec["pass"] == (_rec["throughput"] >= 95.0), _rec
    # >= 边界语义：Throughput 恰为 95.0% 判 PASS；94.9% 判 FAIL
    _edge = _SensScpi(knee=-124.0, tp_pass=95.0)
    _re = measure_sensitivity(_edge, dict(cfg, bler_timeout="3"),
                              {"sensitivity"}, band="OB1", bw="B100",
                              theory_seed=seed12)
    assert _re["result_epre"] is not None, _re["result_epre"]
    _edge_fail = _SensScpi(knee=-124.0, tp_pass=94.9)
    _ref = measure_sensitivity(_edge_fail, dict(cfg, bler_timeout="3"),
                               {"sensitivity"}, band="OB1", bw="B100",
                               theory_seed=seed12)
    assert _ref["result_epre"] is None, _ref["result_epre"]
    # 门限可配置：sens_throughput=90 时 TPut=92% 的点应判 PASS
    _cfg90 = dict(cfg, sens_throughput="90", bler_timeout="3")
    _s90 = _SensScpi(knee=-124.0, tp_pass=92.0)
    _r90 = measure_sensitivity(_s90, _cfg90, {"sensitivity"},
                               band="OB1", bw="B100", theory_seed=seed12)
    assert abs(_r90["result_epre"] - (-124.0)) <= 0.11, _r90["result_epre"]
    assert _r90["criterion"] == "Throughput >= 90.0%", _r90["criterion"]
    # 统计无效（响应 "0,0,,,," 无块计数）：全程无效 + UE 真掉线（PSW=ON）→
    # 首点重测一次仍无效 → 判掉线停止，全程只 1 点（防把 UE 探掉线）。
    _low = _SensScpi(knee=-124.0, invalid_below=1e9, psw_on_at=2)
    _rl = measure_sensitivity(_low, dict(cfg, bler_timeout="3"), {"sensitivity"},
                              band="OB1", bw="B100", theory_seed=seed12)
    assert _rl["result_epre"] is None, _rl["result_epre"]
    assert len(_rl["table"]) == 1, _rl["table"][:3]
    assert _rl["table"][0]["invalid"] is True and _rl["table"][0]["sample_ok"] is False, \
        _rl["table"][0]
    assert _low.inits > len(_rl["table"]), (_low.inits, len(_rl["table"]))
    assert _rl["table"][0].get("dropped") is True, _rl["table"][0]
    # 1dB 退回重降（2026-11-25 真机 085027.log 组合1 早停 2.7dB 根因修复）：
    # coarse 5dB 首个 FAIL 点无统计（UE 被大步跳进失步区）→ 退回上一有效 PASS
    # 改 1dB 步进重降 → UE 始终保持有效统计 → 在 knee 处有效 FAIL 后二分，
    # 探准真实拐点 -124.0（不再塌缩在粗扫点）。
    _iw = _SensScpi(knee=-124.0, invalid_below=-125.5)
    _iw_logs = []
    _rw = measure_sensitivity(_iw, dict(cfg, bler_timeout="3"), {"sensitivity"},
                              band="OB1", bw="B100", theory_seed=seed12,
                              log=_iw_logs.append)
    assert _rw["result_epre"] is not None and abs(_rw["result_epre"] + 124.0) <= 0.11, \
        _rw["result_epre"]
    assert any(r["invalid"] for r in _rw["table"]), _rw["table"][-3:]
    assert min(r["level"] for r in _rw["table"]) >= -126.09, _rw["table"][:3]
    assert any("退回上一有效电平" in x and "1.0dB 步进重降" in x for x in _iw_logs), \
        _iw_logs[-6:]
    # 有效 FAIL 点 + PSW=ON（UE 真掉线）：立即停止二分并标记 dropped，不把
    # 它当弱边界反复探测；结果保留上一有效 PASS 点（-121.08）。
    _dw = _SensScpi(knee=-124.0, psw_on_at=2)
    _rd_logs = []
    _rd = measure_sensitivity(_dw, dict(cfg, bler_timeout="3"), {"sensitivity"},
                              band="OB1", bw="B100", theory_seed=seed12,
                              log=_rd_logs.append)  # 带 log：验证掉线日志格式化不崩
    assert any("UE 已脱网" in x and "判定掉线" in x for x in _rd_logs), _rd_logs[-5:]
    assert _rd["table"][-1].get("dropped") is True, _rd["table"][-3:]
    assert _rd["result_epre"] is not None and abs(_rd["result_epre"] + 121.08) <= 0.2, \
        _rd["result_epre"]
    assert len(_rd["table"]) < len(_rw["table"]), (len(_rd["table"]), len(_rw["table"]))
    # 修正回归（真机 145949.log）：深 FAIL（TPut=0%）但 PSW=ATT（UE 仍在网，
    # 极弱信号收不到 PDCCH）→ 绝不判掉线，继续二分探准拐点 -124.0。
    _wo_logs = []
    _ro = measure_sensitivity(_SensScpi(knee=-124.0),
                              dict(cfg, bler_timeout="3"), {"sensitivity"},
                              band="OB1", bw="B100", theory_seed=seed12,
                              log=_wo_logs.append)
    assert not any(r.get("dropped") for r in _ro["table"]), _ro["table"][-3:]
    assert any("但 UE 仍在网" in x for x in _wo_logs), _wo_logs[-6:]
    assert _ro["result_epre"] is not None and abs(_ro["result_epre"] + 124.0) <= 0.11, \
        _ro["result_epre"]
    # 单点无统计块但重测恢复 → 该点判 PASS，结果仍探到 knee。
    _fl = _SensScpi(knee=-124.0, flaky_q=2)  # 第2次 REL?=首个扫描点首测
    _rf = measure_sensitivity(_fl, dict(cfg, bler_timeout="3"), {"sensitivity"},
                              band="OB1", bw="B100", theory_seed=seed12)
    assert _rf["result_epre"] is not None and abs(_rf["result_epre"] + 124.0) <= 0.11, \
        _rf["result_epre"]
    # 重测后该点恢复有效统计（INIT 次数比扫描点数多 1 次重测）
    assert _fl.inits > len(_rf["table"]), (_fl.inits, len(_rf["table"]))
    assert all(r["sample_ok"] for r in _rf["table"]), _rf["table"][:2]
    # 扫描收尾流程（用户确认 + 真机 151333.log）：先拉回强信号，纯灵敏度不切 RMC，
    # 含 TX 指标才切 RMC；杜绝在弱信号下切 STYP 把在网 UE 搞掉线→重建小区白耗 15s。
    class _TailCmd(_SensScpi):
        def __init__(self, *a, **kw):
            super().__init__(*a, **kw)
            self.cmds = []
        def write(self, cmd):
            self.cmds.append(cmd)
            super().write(cmd)
    _tl = _TailCmd(knee=-124.0)
    measure_sensitivity(_tl, dict(cfg, bler_timeout="3"), {"sensitivity"},
                        band="OB5", bw="B014", theory_seed=seed12)
    _idx_tail = max(i for i, c in enumerate(_tl.cmds) if "RSEPre:LEVel -85.0" in c)
    # 收尾拉强信号后不得再有 STYP/RMC 配置（纯灵敏度场景）
    assert not any("STYP RMC" in c or "RMC:UL" in c
                   for c in _tl.cmds[_idx_tail + 1:]), _tl.cmds[_idx_tail + 1:]
    assert _tl.levels[-1] == -85.0, _tl.levels[-3:]
    _tltx = _TailCmd(knee=-124.0)
    measure_sensitivity(_tltx, dict(cfg, bler_timeout="3"),
                        {"sensitivity", "maxpower"},
                        band="OB5", bw="B014", theory_seed=seed12)
    _idx_tailtx = max(i for i, c in enumerate(_tltx.cmds) if "RSEPre:LEVel -85.0" in c)
    assert any("STYP RMC" in c for c in _tltx.cmds[_idx_tailtx + 1:]), \
        "含 TX 指标时收尾应切回 RMC"
    # 二分精扫精度 + 极端偏差兜底：理论种子 -124，实测 knee 偏离 ±8dB 仍要探准
    for _knee in (-116.0, -132.0, -134.5):
        _s2 = _SensScpi(knee=_knee)
        _r2 = measure_sensitivity(_s2, dict(cfg, bler_timeout="3"), {"sensitivity"},
                                  band="OB1", bw="B100", theory_seed=seed12)
        assert abs(_r2["result_epre"] - _knee) <= 0.11, (_knee, _r2["result_epre"])
    # --- TX/RX 分模块 + 灵敏度功率档 + 动态列 ---
    assert TX_METRIC_KEYS == ["maxpower", "aclr", "evm", "freqerr"]
    assert RX_METRIC_KEYS == ["sensitivity"]
    # 功率档命令：high=MAXPower；low 有值=CLOop+CLTPower；low 空=MINPower
    hi_cmd, hi_desc = build_sens_power_cmd(cfg, "high", -10.0)
    lo_cmd, lo_desc = build_sens_power_cmd(cfg, "low", -10.0)
    lo_min_cmd, lo_min_desc = build_sens_power_cmd(cfg, "low", None)
    assert hi_cmd == ["CONF:LTE:SIGN:UL:TPC:SET MAXPower"], hi_cmd
    assert any("CLTPower -10.0" in c for c in lo_cmd) and \
        any("SET CLOop" in c for c in lo_cmd), lo_cmd
    assert lo_min_cmd == ["CONF:LTE:SIGN:UL:TPC:SET MINPower"], lo_min_cmd
    assert "MAXPower" in hi_desc and "CLOop" in lo_desc and "MINPower" in lo_min_desc
    # 动态列：只勾灵敏度（无 TX 指标）→ 电压/电流列隐藏，UL RB+sens Limit 在灵敏度之前；
    # both → Sens@Pmax/Sens@Pmin + Desense
    only_sens = {"maxpower": False, "aclr": False, "evm": False,
                 "freqerr": False, "sensitivity": True}
    h_sens = wide_headers(only_sens, "high", True)
    assert h_sens == ["Band", "BW", "DL Channel", "LCRB @ RBstart",
                      "Sens limit\n(dBm)", "Sens@Pmax\n(dBm)"], h_sens
    h_both = wide_headers(only_sens, "both", True)
    assert h_both[-5:] == ["LCRB @ RBstart", "Sens limit\n(dBm)",
                           "Sens@Pmax\n(dBm)", "Sens@Pmin\n(dBm)", "Desense\n(dB)"], h_both
    h_low = wide_headers(only_sens, "low", True)
    assert h_low[-1] == "Sens@Pmin\n(dBm)" and len(h_low) == len(h_sens), h_low
    # 电源关闭时 Voltage/Current 列消失；勾了 TX 指标时仍显示
    h_nopower = wide_headers(only_sens, "high", False)
    assert "Voltage\n(V)" not in h_nopower and "Current@Pmax\n(mA)" not in h_nopower
    h_tx = wide_headers(dict(only_sens, maxpower=True), "high", True)
    assert "Voltage\n(V)" in h_tx and "Current@Pmax\n(mA)" in h_tx
    # 行与动态表头对齐 + 双灵敏度列 + Desense 取值
    d2 = dict(demo_measurement(dict(cfg, metrics=only_sens)))
    d2["sensitivity_low"] = {"result": round(-125.8 + full_cell_bw_offset_db("B100"), 2)}
    wide_both = make_wide_row("B1 / 10MHz / 低(CH50)", d2, metrics=only_sens,
                              sens_power_mode="both", psu_enable=True)
    assert len(wide_both) == len(h_both), (len(wide_both), len(h_both))
    assert wide_both[-3] != wide_both[-2]  # 大/小功率两值不同
    # Desense = 大功率灵敏度 - 小功率灵敏度
    assert abs(float(wide_both[-1]) -
               (float(wide_both[-3]) - float(wide_both[-2]))) <= 0.11, wide_both
    # 电源命令
    pv = build_psu_voltage_commands(dict(cfg, psu_enable=True, psu_voltage="3.8", psu_channel="P6V"))
    assert "SOUR:VOLT 3.8000,(@1)" in pv and "OUTP ON,(@1)" in pv
    # 结果行
    d = demo_measurement(dict(cfg, metrics={k: True for k in METRIC_KEYS}))
    rows = make_rows(d, {"evm_lim": "17.5", "aclr_lim": "30", "freq_lim": "100"})
    assert len(rows) >= 12
    assert any(r[1] == "Power (TX)" for r in rows)
    assert "current" not in METRIC_KEYS  # 电流改为默认记录，不再是勾选项
    assert any(r[1] == "灵敏度 (Full Cell BW)" for r in rows)
    # Full Cell BW 换算：10MHz 偏置 10log10(600)=27.78dB；-110 RS EPRE -> -82.22dBm
    assert abs(full_cell_bw_offset_db("B100") - 27.78) < 0.02
    assert abs(epre_to_fullbw(-110.0, "B100") - (-82.22)) < 0.02
    assert abs(epre_to_fullbw(-124.8, "B100") - (-97.02)) < 0.02
    assert abs(full_cell_bw_offset_db("B014") - 18.57) < 0.02  # 12*6=72
    # 界面显示名 / 内部 SCPI 键映射
    assert band_label("OB1") == "B1" and band_label("OB111") == "B111"
    assert bw_label("B014") == "1.4MHz" and bw_label("B100") == "10MHz"
    assert bw_label("B030") == "3MHz" and bw_label("B200") == "20MHz"
    assert bw_code(1.4) == "B014" and bw_code(10) == "B100"
    assert combo_display("OB1", "B100") == "B1 / 10MHz"
    # 数据全部来自本目录 specs\LTE\bands.json（不再依赖任何外部路径）
    assert len(ALL_BANDS) >= 60, len(ALL_BANDS)
    assert sum(len(v) for v in channel_table().values()) >= 200
    _src = channel_table_source()
    assert ("bands.json" in _src or "lte_channels" in _src), _src
    assert ("LTE" + "&NR") not in _src, _src
    assert os.path.isfile(LOCAL_BANDS_JSON), LOCAL_BANDS_JSON
    assert band_supported_bws("OB1") == ["B050", "B100", "B150", "B200"]
    # 勾选频段的默认带宽：只勾 10MHz；不支持 10MHz 的频段取最接近的支持带宽
    assert default_bw_for_band("OB1") == "B100"
    assert default_bw_for_band("OB2") == "B100"
    assert default_bw_for_band("OB31") == "B050"    # 1.4/3/5MHz → 5MHz
    assert default_bw_for_band("OB106") == "B030"   # 1.4/3MHz → 3MHz
    assert all(default_bw_for_band(b) in band_supported_bws(b)
               for b in ALL_BANDS if band_supported_bws(b))
    # 频段×带宽低/中/高信道（本地 R19 快照，按带宽内缩 B/2）
    ch = band_channels("OB1", "B100")
    assert ch == [("低", 50), ("中", 300), ("高", 550)], ch
    assert band_channels("OB1", "B014") == []          # B1 不支持 1.4MHz → 跳过
    assert band_channels("OB12", "B100")[0][1] == 5060
    assert band_channels("OB22", "B200")[-1][1] == 7300
    assert band_channels("OB6", "B100") == [("低", 2700)]  # 窄频段三点重合 → 去重
    assert "CONF:LTE:SIGN:RFS:CHAN:DL 300" == build_channel_cmd(cfg, 300)
    # 每个频段独立带宽：旧全局 bws 只迁移给旧 bands，不扩散到其它频段
    nb = normalize_band_bws({"bands": ["OB1", "OB2"], "bws": ["B100"],
                             "band_bws": {}})
    assert nb["OB1"] == ["B100"] and nb["OB2"] == ["B100"], nb
    assert nb["OB6"] == [], nb
    # 不支持的带宽在规范化时被过滤（B1 不支持 1.4MHz）
    nb2 = normalize_band_bws({"band_bws": {"OB1": ["B014", "B100"]}})
    assert nb2["OB1"] == ["B100"], nb2
    # 测试项 = 各频段自己的带宽，B1 只测 10MHz、B2 测 1.4/20MHz，无笛卡尔积
    sel = {"OB1": ["B100"], "OB2": ["B014", "B200"]}
    cmb, _skp = build_combos(["OB1", "OB2"], sel)
    assert len(cmb) == 9, cmb
    assert ("OB1", "B100", "低", 50) in cmb
    assert all(c[1] == "B100" for c in cmb if c[0] == "OB1")
    assert all(c[1] in ("B014", "B200") for c in cmb if c[0] == "OB2")
    assert not any(c[0] == "OB1" and c[1] == "B014" for c in cmb)
    assert "B1: 10MHz" in band_bws_summary(["OB1", "OB2"], sel), \
        band_bws_summary(["OB1", "OB2"], sel)
    # 全信道扫描：B1/10MHz 合法 DL EARFCN = 50..550（501 个），首尾与三点表一致
    _all = all_band_channels("OB1", "B100")
    assert len(_all) == 501 and _all[0] == ("全", 50) and _all[-1] == ("全", 550), \
        (_all[:3], _all[-3:], len(_all))
    _cmb_full, _skp_full = build_combos(["OB1"], {"OB1": ["B100"]}, full_scan=True)
    assert len(_cmb_full) == 501 and _cmb_full[0][3] == 50 and _cmb_full[-1][3] == 550
    assert CONFIG_DEFAULTS["full_scan"] is False
    # 宽表：一行一个组合，指标按列（UL RB+sens_limit 已移到 TX 指标后、灵敏度前）
    _metrics_all = {"sensitivity": True, "maxpower": True, "aclr": True, "evm": True,
                   "freqerr": True, "psu_enable": True}
    wide = make_wide_row("B1 / 10MHz / 中(CH300)", d,
                         {"evm_lim": "17.5", "aclr_lim": "30", "freq_lim": "100"},
                         metrics=_metrics_all)
    hdr_all = wide_headers(_metrics_all)
    assert len(wide) == len(hdr_all)
    assert wide[0] == "B1" and wide[1] == "10MHz" and wide[2] == "CH300", wide[:3]
    # UL RB + sens_limit 已移到 TX 指标之后，按表头名称定位
    assert wide[hdr_all.index("LCRB @ RBstart")] == "50@0"
    assert wide[hdr_all.index("Sens limit\n(dBm)")] == "-96.3"
    # 界面表头：场景标注与单位同行（共 2 行）；内部列名/导出不受影响
    assert header_display("Power@Pmax\n(dBm)") == "Power\n@Pmax (dBm)"
    assert header_display("Current@Pmin\n(mA)") == "Current\n@Pmin (mA)"
    assert header_display("Sens@Pmax\n(dBm)") == "Sens\n@Pmax (dBm)"
    assert header_display("Power@Pmax") == "Power\n@Pmax"
    # 本身含 @ 的普通列名不误伤
    assert header_display("LCRB @ RBstart") == "LCRB @ RBstart"
    assert header_display("Sens limit\n(dBm)") == "Sens limit\n(dBm)"
    assert all(header_display(h) == h for h in ("Band", "BW", "DL Channel"))
    hdr = wide_headers(_metrics_all)
    assert hdr[:6] == ["Band", "BW", "DL Channel", "Voltage\n(V)", "Current@Pmax\n(mA)", "Power@Pmax\n(dBm)"], hdr[:6]
    assert "LCRB @ RBstart" in hdr and "Sens limit\n(dBm)" in hdr and "Sens@Pmax\n(dBm)" in hdr
    assert "Power(W)" not in hdr and "EVM Peak(%)" not in hdr
    assert "22.8" in wide[hdr.index("Power@Pmax\n(dBm)")]  # 按表头定位，不写死列号
    # 公用测试场景：Pmin 单档 TX 列名带 @Pmin；both 双档 TX 列成对出现且行对齐
    h_pmin = wide_headers(_metrics_all, "low", True)
    assert "Power@Pmin\n(dBm)" in h_pmin and "Power@Pmax\n(dBm)" not in h_pmin, h_pmin
    h_both_tx = wide_headers(_metrics_all, "both", True)
    assert "Power@Pmax\n(dBm)" in h_both_tx and "Power@Pmin\n(dBm)" in h_both_tx
    assert "EVM RMS@Pmax\n(%)" in h_both_tx and "EVM RMS@Pmin\n(%)" in h_both_tx
    wide_both_tx = make_wide_row("B1 / 10MHz / 中(CH300)", d,
                                 {"evm_lim": "17.5", "aclr_lim": "30", "freq_lim": "100"},
                                 metrics=_metrics_all, sens_power_mode="both")
    assert len(wide_both_tx) == len(h_both_tx), (len(wide_both_tx), len(h_both_tx))
    assert wide_both_tx[h_both_tx.index("Power@Pmax\n(dBm)")] == \
           wide[hdr.index("Power@Pmax\n(dBm)")], "Pmax 档 TX 值应与单档一致"
    # both 场景若缺 Pmin 档实测数据（mod_avg_low），@Pmin 列必须留空，
    # 不得回退复制 Pmax 的数值冒充（否则两列看起来一样，实际只测了一档）
    assert wide_both_tx[h_both_tx.index("Power@Pmin\n(dBm)")] == "", \
        wide_both_tx[h_both_tx.index("Power@Pmin\n(dBm)")]
    assert wide_both_tx[h_both_tx.index("EVM RMS@Pmin\n(%)")] == ""
    # 真有 Pmin 档数据时两列各自独立取值
    _d_lo = dict(d)
    _d_lo["mod_avg_low"] = list(d["mod_avg"])
    _d_lo["aclr_low"] = list(d["aclr"])
    _d_lo["mod_avg_low"][17] = 5.1
    _w_lo = make_wide_row("B1 / 10MHz / 中(CH300)", _d_lo, None,
                          metrics=_metrics_all, sens_power_mode="both")
    assert _w_lo[h_both_tx.index("Power@Pmax\n(dBm)")] == "22.80", _w_lo
    assert _w_lo[h_both_tx.index("Power@Pmin\n(dBm)")] == "5.10", _w_lo
    # MEV 结果有效性：全 INV 应判 valid=False（UE 未注册场景）
    class _FakeScpi:
        def __init__(self, resp):
            self.resp = resp
            self.writes = []
        def write(self, _cmd):
            self.writes.append(_cmd)
        def query(self, cmd, timeout=None, expect=None):
            if cmd.endswith("MEV:STAT?"):
                return "RDY"
            if "MOD:AVER?" in cmd:
                return self.resp[0]
            if "MOD:EXTR?" in cmd:
                return self.resp[0]
            if "ACLR?" in cmd:
                return self.resp[1]
            return ""
    # 自检时用 0 间隔、0 稳定、2 次重试，避免真实 sleep（真机默认 gap=1.0s / tpc_settle=4s）
    cfg_fast = dict(cfg, mev_retry="2", mev_retry_gap="0", tpc_settle="0")
    bad_scpi = _FakeScpi(("4," + ",".join(["INV"] * 29),
                          "4," + ",".join(["INV"] * 8)))
    bad = measure_tx(bad_scpi, cfg_fast, {"maxpower", "evm", "aclr"})
    assert bad["valid"] is False, bad["valid"]
    assert bad["reliability"] == 4, bad["reliability"]        # 首位 Reliability=Underdriven
    assert "Underdriven" in bad["reliability_text"]
    assert bad["diagnosis"] and "PUSCH" in bad["diagnosis"]
    assert bad["attempts"] == 3, bad["attempts"]              # 首次 + 2 次重试
    good_scpi = _FakeScpi((d["raw"]["mod_avg"], d["raw"]["aclr"]))
    good = measure_tx(good_scpi, cfg_fast, {"maxpower", "evm", "aclr"})
    assert good["valid"] is True, good["valid"]
    assert good["reliability"] == 0, good["reliability"]
    assert good["attempts"] == 1, good["attempts"]            # 有效即返回，不重试
    # 重试时重发 TPC MAXPower（2 次重试 → 2 条），而非仅重新 INIT
    assert bad_scpi.writes.count("CONF:LTE:SIGN:UL:TPC:SET MAXPower") == 2, \
        bad_scpi.writes
    # Pmin 场景重试必须重发 low 档功率命令（CLOop + CLTPower），不能回 Pmax
    low_bad_scpi = _FakeScpi((("4," + ",".join(["INV"] * 29)),
                              ("4," + ",".join(["INV"] * 8))))
    low_bad = measure_tx(low_bad_scpi, dict(cfg_fast, sens_low_power="-20"),
                         {"maxpower"}, power_mode="low")
    assert low_bad["valid"] is False
    assert low_bad_scpi.writes.count("CONF:LTE:SIGN:UPLink:TPC:CLTPower -20.0") == 2, \
        low_bad_scpi.writes
    assert "CONF:LTE:SIGN:UL:TPC:SET CLOop" in low_bad_scpi.writes
    assert "CONF:LTE:SIGN:UL:TPC:SET MAXPower" not in low_bad_scpi.writes
    # Rel=3 Overdriven 但关键字段有效 → 接受结果并标记，不重试/中止
    ov_scpi = _FakeScpi(("3," + d["raw"]["mod_avg"].split(",", 1)[1], d["raw"]["aclr"]))
    ov = measure_tx(ov_scpi, cfg_fast, {"maxpower", "evm", "aclr"})
    assert ov["valid"] is True and ov["reliability"] == 3, ov
    assert ov["overdriven_but_usable"] is True and ov["attempts"] == 1, ov
    # 导出
    import tempfile
    td = tempfile.mkdtemp(prefix="cmw_self_")
    result = {"time": "2026-01-01T00:00:00", "cfg": cfg,
              "items": [{"combo": "B2 / 10MHz / 低(CH650)", "data": d,
                         "rows": rows, "wide": wide}]}
    ok, err = export_results(result, os.path.join(td, "r.xlsx"), os.path.join(td, "r.txt"))
    assert ok, err
    print("selftest OK | 命令校验通过 | 结果行=%d | 导出=%s" % (len(rows), "OK" if ok else err))


def run_demo():
    cfg = dict(CONFIG_DEFAULTS)
    cfg["metrics"] = {k: True for k in METRIC_KEYS}
    cfg["bands"] = ["OB2"]
    cfg["band_bws"] = {"OB2": ["B100"]}
    cfg["bws"] = ["B100"]
    data = demo_measurement(cfg)
    limits = {"evm_lim": "17.5", "aclr_lim": "30", "freq_lim": "100"}
    print("演示结果（B2 / 10MHz，低/中/高三信道，指标按列）")
    headers = wide_headers()
    print(" | ".join(headers))
    td = os.path.join(RESULT_DIR, "demo")
    os.makedirs(td, exist_ok=True)
    items = []
    for label, _ch in [(("B2 / 10MHz / %s(CH%d)" % (p, e)), e)
                       for p, e in band_channels("OB2", "B100")]:
        wide = make_wide_row(label, data, limits)
        print(" | ".join(wide))
        rows = make_rows(data, limits)
        items.append({"combo": label, "data": data, "rows": rows, "wide": wide})
    result = {"time": datetime.datetime.now().isoformat(timespec="seconds"),
              "cfg": cfg, "items": items}
    ok, err = export_results(
        result, os.path.join(td, "demo.xlsx"), os.path.join(td, "demo.txt"))
    print("\n示例报告已生成: %s (%s)" % (td, err or "xlsx+txt"))


# ----------------------------------------------------------------------------
# GUI
# ----------------------------------------------------------------------------

HELP_TEXT = r"""使用要点（依据 CMW500 Operating Manual 2.0.10 / E36300 编程指南）

1. 连接
   - 网线接 CMW500 后面板 LAN(REMOTE) 口，Data Port 默认 5025。
   - 电源 E36312A 同样走 LAN socket 5025、LF 结尾，复用同一套传输。
   - IP/端口/RF口/电压等全部自动保存到 config.json，下次启动自动恢复。

1a. AT 控制（USB 串口 / 手动 debug）
   - 位置：CMW500 配置行右侧「AT 控制」框。AT口 下拉框默认 auto，启动时
     自动枚举 COM 口并逐个 AT 握手（描述含 modem/asr 的优先），命中的口
     自动选中；也可手动点【检测】重探或直接填 COMx。
   - 用法：指令框输入 AT 指令（如 AT+CFUN?、AT+CFUN=0、AT+CFUN=1）后点
     【发送】或回车。支持多条指令依次执行，用 ; 分隔（如
     AT+CFUN=0; AT+CFUN=1），同一串口会话内逐条下发，单条失败不中断后续。
     后台线程执行，日志实时回显（[AT] COMx > 指令 / [AT] COMx < 响应），
     读到 OK/ERROR 或 8s 超时为止，不阻塞测量。
   - 输入框内容自动保存（关闭工具/每次发送后写入 config.json），下次启动恢复。
   - 自动恢复：RX 测试掉线重注册时（硬掉线/快切未回网），若 AT 口可用会自动
     下发 CFUN=0→1 强制 UE 重注册（强信号下秒级回网），未生效再回退重建小区；
     config.json 的 at_auto_recover=0 可禁用，at_recover_wait 控制回网等待。
   - 典型 debug：CFUN=0（关 RF）→ CFUN=1（重新搜网注册仪表），强制 UE
     主动重注册，替代被动等自恢复；CEREG?/CSQ 查看 UE 侧注册状态与信号。
   - 串口按次打开关闭（115200），指令执行期间【发送】按钮禁用防重入。

2. 测试流程（信令模式）
   ① 填好 CMW500 与电源两行配置；在「工作频段与带宽」表里勾选多个频段，
      再逐行勾选各频段要测的带宽，或直接点表头带宽复选框整列一次生效
      （灰色 — = 该频段不支持此带宽）；最后勾选要测的指标；
   ② 点【初始化仪表】：连接两台仪器 → 复位 → 配置 LTE 信令小区
      (BAND/带宽/DL功率/UL PMAX/RMC) → 开小区 → 让 DUT 附着建连接
      → 电源下发电压并开启输出；
   ③ 先选【测试场景】：Pmax / Pmin / Pmax+Pmin（TX/RX 公用），再勾选
      TX 指标（功率/ACLR/EVM/频率误差）与 RX 指标（灵敏度）；
   ④ 点【开始测量】：按「每个已选频段各自勾选的带宽」从 R19 信道表取
      低/中/高三个 DL EARFCN
      （已按带宽内缩 B/2，见对照表），逐组合按所选场景设置 UE 上行
      功率档，再用 CSPath 测该场景的 TX 指标；灵敏度探底（下行功率从高到低
      扫描，找 Throughput≥95% 的最低电平）也在同一场景下进行；
   ⑤ 测完弹窗提示，并自动在 result/ 目录生成 时间戳.xlsx 和 .txt。
      报告为宽表：一行一个频段/带宽/信道，只列勾选的指标；
      「Pmax+Pmin」时 TX 指标拆成 @Pmax/@Pmin 双列，灵敏度两列 + Desense。
   ★ 暂停/继续：测量中点【暂停】，线程在当前组合测完后挂起（已测数据全部
     保留、界面保持不动）；再点【继续】从下一组合接着测。暂停粒度为一个
     完整组合，不打断单个指标/灵敏度的 SCPI 测量序列。

3. 指标说明
   - 功率 / ACLR / EVM / 频率误差：连接态下 Multi Evaluation
     测 DUT 上行（FETC:LTE:MEAS:MEV，手册 21.5.3），在所选测试场景的
     功率档下测得（Pmax 场景即 TPC MAXPower 满功率）。
   - 电压 / 电流：电源启用且勾选了 TX 指标时默认同步记录，无需勾选；E36312A 读
     MEAS:CURR?/MEAS:VOLT?（测量前读待机、测量时读发射）。只测 RX 灵敏度
     （未勾任何 TX 指标）时，电压/电流列不显示。
     结果表组合列 = 频段 / 带宽 / DL Channel（工具实际下发的是 DL EARFCN，
     即 CONF:LTE:SIGN:RFS:CHAN:DL；TDD 频段 UL/DL 同频，仍按 DL 口径）/
     UL RB配置（LCRB@RBstart，取自本地
     ul_cfg.json，3GPP TS 36.521-1 表 7.3.5-2；无数据显 '—'）/
     3GPP灵敏度Limit(dBm)（本地 refsens.json 的 REFSENS，FullBW 口径），固定显示；
     其后只列勾选的指标（TX：电压 → 电流 → 功率 → ACLR → EVM RMS → 频率误差；
     RX：灵敏度），未勾选的指标列不显示；TX 列名带 @Pmax/@Pmin 场景标注，
     「Pmax+Pmin」时 TX 拆双列，末尾为大功率灵敏度 / 小功率灵敏度 /
     Desense(=大功率-小功率) 三列。
     电源功率(W) 与 EVM Peak 不再显示（原始值仍在内部明细里备查）。
   - 灵敏度(探底)：CONF:LTE:SIGN:DL:RSEPre:LEVel 扫描 + INIT/FETC:EBLer
     （手册 23.4.2 / 23.5.10）。判定量 = FETC:LTE:SIGN:EBLer:REL? 第 5 值
     Throughput(%)（idx 4），Throughput ≥ 95%（config 的 sens_throughput 可改）
     判 PASS，与 ASRrun 参考工具/3GPP 36.521-1 判据一致。
     单点单次查询，不再查 RELative?（减少 SCPI 往返、避免响应错位）；
     明细页记录每点 Throughput 供核对。
     ★ 功率口径：仪器设置的是 RS EPRE(dBm/15kHz)，报告的灵敏度是
       Full Cell BW Pow(dBm)，按手册 23.3.6 换算 FullBW=RS+10log10(12*N_RB)
       （10MHz 即 +27.78dB；-110 RS EPRE 仅 -82.2dBm，远非灵敏度）。
     ★ 扫描策略：三级扫描 粗步(5dB)→中步(1dB)→精步(0.1dBm)，越拐点逐级回退
       精扫，精度 0.1dBm，兼顾速度与精度。首个信道用该频段/带宽的 3GPP 理论
       REFSENS(实测误差<1dB)作种子，从理论值+3dB 起步，跳过 -85 起的强信号
       空扫（省约 13s）；理论偏差大时仍能粗扫到全局下限，不会探不到底。
       同频段同带宽后续信道从上一信道灵敏度上方(默认+3dB)窄窗口续扫，
       不重复全窗（三信道实测差异通常 <1dB）。
     ★ 扫描下限默认 -135dBm RS EPRE（10MHz REFSENS 约 -97dBm FullBW，
       对应 RS EPRE 约 -125），高级参数可改 coarse/near/margin/span。
     ★ 低带宽灵敏度受 CMW 低电平输出底噪限制（真机 2026-11-23 B5 实证：
       1.4MHz 与 3MHz 的 FullBW 灵敏度几乎相同——拐点总功率恒定约 -101dBm，
       1.4MHz 比理论差约 4dB，均为 CMW 输出底噪钳位，非 UE 真实灵敏度）。
       需测 1.4/3MHz 真实灵敏度时：CMW 输出口串接固定衰减器（建议 30~40dB），
       并在 RX 指标框填「外部衰减(dB)」——初始化时经 EATT:OUTP/EATT:INP 成对
       补偿（手册 p.1729），DL/UL 电平自动换算，报告仍为 UE 天线口电平。
     灵敏度前按 3GPP TS 36.521-1 表 7.3.5-2 切 UDCH 精确配置参考信道 RB：
     UL = LCRB@RBstart（逐 band/带宽查 specs/LTE/ul_cfg.json），DL = 满 RB@0；
     测完自动切回 RMC。切 STYP 会使 UE 掉线，工具会主动发 CALL:PSW:ACT CONN
     触发重新附着（被动等待无效，真机已验证）。
     ★ 测试场景（TX/RX 公用，UE 上行发射功率；小功率会影响灵敏度 desense）：
       大功率 = TPC:SET MAXPower（UE 升到最大功率，约 23dBm）；
       小功率 = UPLink:TPC:CLTPower <P> + UL:TPC:SET CLOop（闭环功率，
                 默认 -10dBm，范围 -50~23dBm，见手册 p.1711/1752/1753）；
       大+小功率 = 每信道按两档各测一遍，分别缓存热点，结果表出两列灵敏度
                   （大功率灵敏度FullBW / 小功率灵敏度FullBW），并在末尾加一列
                   Desense(dB) = 大功率灵敏度 - 小功率灵敏度。
       切功率档前同样把电平拉回强信号，避免弱信号切换掉线。
     注意：UDCH 需 CMW-KS510 选件；无此选件时仪表会拒绝 STYP UDCH。
   - 频段界面显示 B1/B2/…，带宽显示 1.4MHz/3MHz/5MHz/10MHz/15MHz/20MHz；
     下发给仪表的 SCPI 仍用 CONF:LTE:SIGN:BAND OB<n> 与
     CONF:LTE:SIGN:CELL:BAND:DL B0xx（CMW500 固定语法，不能改）。
   - 频段清单、各频段支持的带宽、低/中/高信道全部取自本目录
     specs\LTE\bands.json（RF Band Tool / 3GPP R19 的本地快照，随工具一同复制）。
     工具运行与刷新只读本目录，不访问任何外部路径；该文件缺失时回退
     lte_channels.json 全量快照，再回退内置兜底表。
   - 全信道扫描（RX 指标框复选框）：默认关，每个 (频段,带宽) 只测低/中/高三点；
     勾选后按 TS 36.101 口径枚举该带宽下全部合法 DL EARFCN（同信道表窗口，
     离频段边缘 ≥ B/2），逐信道灵敏度探底。例如 B1/10MHz = CH50~550 共 501 点，
     耗时约为三信道的 167 倍，仅用于全面摸底。结果表首行与左三列在横/纵滚动时固定。
     运行 python refresh_lte_channels.py 可刷新快照与对照表。
     鼠标悬停频段复选框可查看该频段支持的带宽与三个信道。
   - 每个频段可独立选择不同带宽（例如 B1 只测 10MHz，B2 测 1.4/5/20MHz），
     测试项 = 各频段自己勾选的带宽，不再做全局 频段×带宽 笛卡尔积。
     FDD+TDD 全部列出；按 频段×该频段带宽×低/中/高三信道 依次测量，
     报告一行一个组合、指标按列。
   - 【批量】先勾选多个频段，再点表头某个带宽的复选框，即可让这些频段
     一次性全部选上（或全部取消）该带宽，无需逐行点选；只作用于已勾选
     且支持该带宽的频段，灰色表头表示当前没有可作用的频段。
     两个按钮可批量「已选频段全选带宽 / 清空全部选择」；后者会同时
     取消已勾选的频段与全部带宽，一键恢复空白选择状态。
   - 勾选某频段时默认【只勾 10MHz】（不再是自动全选所有支持带宽）；
     不支持 10MHz 的频段（B31/B51/B54/B72/B73/B87/B88/B106）默认勾
     最接近的支持带宽（5MHz 或 3MHz）并在状态栏提示。
     已手动勾过带宽的频段保持原样；想恢复默认可先点「清空全部选择」再重勾，
     「已选频段全选带宽」按钮仍可一次全选全部支持带宽。
     表格可滚动，鼠标悬停频段名可查看支持的带宽与三个信道；
     顶部汇总实时显示 已选频段数 / 频段-带宽组合数 / 实际测试点数。
   - 低/中/高信道按当前带宽内缩 B/2（TS 36.101 5.7.3），保证整条
     信道落在频段内，不能取频段范围两端：
       例 B1 / 10MHz = CH50 / CH300 / CH550，而非 CH0 / CH299 / CH599。
     仅测该频段实际支持的带宽，不支持的组合自动跳过并在日志注明；
     窄频段下三点重合时去重（只测一次）。UL 信道由仪表自动换算。
     完整对照表见《LTE频段_带宽_高中低信道对照表.txt》。
   - 高级参数不再显示在界面上，沿用 config.json / 默认值。

3.1 上行测量无有效结果时（弹窗/日志会给出 Reliability 诊断）
   - FETC 返回数组首位是 Reliability Indicator（手册 p.1736）：
       0=OK | 1=测量超时 | 3=输入过强 | 4=输入上行信号过低(Underdriven)
       6=无触发 | 7=无法解码 | 8=无法同步 | 15=参考频率错误。
   - 若为 4(Underdriven)：CMW 没收到足够强的 UE 上行信号，重点查
       ① UE 是否在 PUSCH 持续发射（电流是否升到发射态）；
       ② 期望标称功率 RFS:ENPower 是否远高于 UE 实际功率（默认 23dBm，
          过高会使输入量程失配而判 Underdriven）；
       ③ 射频连线 / 外部衰减 EATT / RF 口是否与 DUT 一致；
       ④ Uplink Power Control：PUSCh:LEVel / PMAX / TPC MAXPower；
       ⑤ 下行 RS EPRE 是否足够（默认 -80dBm）。
   - 切换信道（同频段同带宽）后连接进入 Signaling in Progress 重同步
     （手册 p.1681）：工具等待 chan_settle，并对首个全 INV 组合重新触发
     测量重试 mev_retry 次。
   - 切换频段/带宽则不同：连接态直接改 BAND 会让 UE 掉线且仪表忙 ~10s
     不响应查询（真机 log 已复现）。工具先尝试快切：改参数+幂等 CELL ON，
     事件驱动等待（小区 ON,ADJ → UE 回网窗口）——UE 跟上即省去重建；
     未跟上则 AT CFUN=0→1 强制重注册（更快更确定），AT 也无效才走
     「关小区→重配→开小区→等 ON,ADJ→CALL:PSW:ACT CONN→等 UE 重新
     ATTACHED」完整重建（手册 p.1732/1733）。
   - 仍全部无效即中止，不生成误导性空报告（手册 p.1680 排查清单）。

4. 手册章节
   - PDF p.139/2671       远程控制 / Direct Socket
   - PDF p.1522~1524      信号路由（CSPath 联合信号路径）
   - PDF p.1528           21.5.3 Multi Evaluation 命令
   - PDF p.1728/1734      23.4.1 小区配置 / 23.4.2 BLER 测试
   - PDF p.1737/1768      23.5.2 连接控制 / 23.5.10 BLER 命令
   - E36311-90008 p.5/17/69  SOUR:VOLT / 通道 / MEAS:CURR?
"""


def _attach_band_tip(widget, band):
    """频段复选框悬停提示：该频段支持的带宽及低/中/高信道（本目录 data\\LTE\\bands.json）。"""
    rec = channel_table().get(band) or {}
    lines = ["频段 %s  支持带宽 / 低·中·高 DL EARFCN:" % band_label(band)]
    for w in BW_CHOICES:
        h = rec.get(w)
        if h:
            lines.append("  %-7s 低 CH%-6d 中 CH%-6d 高 CH%-6d" %
                         (bw_label(w), h[0], h[1], h[2]))
    if len(lines) == 1:
        lines.append("  （该频段无可用带宽）")
    lines.append("数据源: 本目录 data\\LTE\\bands.json  3GPP R19 本地快照")
    text = "\n".join(lines)
    state = {"win": None}

    def show(_e=None):
        if state["win"]:
            return
        win = tk.Toplevel(widget)
        win.wm_overrideredirect(True)
        win.wm_geometry("+%d+%d" % (widget.winfo_rootx() + 16,
                                    widget.winfo_rooty() + widget.winfo_height() + 2))
        ttk.Label(win, text=text, background="#ffffe0", relief="solid",
                  borderwidth=1, padding=(8, 5), justify="left").pack()
        state["win"] = win

    def hide(_e=None):
        if state["win"]:
            state["win"].destroy()
            state["win"] = None

    widget.bind("<Enter>", show)
    widget.bind("<Leave>", hide)


def _attach_full_scan_tip(widget):
    """全信道扫描复选框悬停提示：说明范围与耗时。"""
    text = ("全信道扫描（RX 灵敏度探底）\n"
            "勾选后：每个已勾选频段×带宽，按 3GPP TS 36.101 口径枚举\n"
            "该带宽下全部合法 DL EARFCN（含低/中/高之间所有信道），\n"
            "逐信道执行灵敏度探底。信道数远多于默认 3 点，耗时显著增加。\n"
            "未勾选：保持默认低/中/高三信道。")
    state = {"win": None}

    def show(_e=None):
        if state["win"]:
            return
        win = tk.Toplevel(widget)
        win.wm_overrideredirect(True)
        win.wm_geometry("+%d+%d" % (widget.winfo_rootx() + 16,
                                    widget.winfo_rooty() + widget.winfo_height() + 2))
        ttk.Label(win, text=text, background="#ffffe0", relief="solid",
                  borderwidth=1, padding=(8, 5), justify="left").pack()
        state["win"] = win

    def hide(_e=None):
        if state["win"]:
            state["win"].destroy()
            state["win"] = None

    widget.bind("<Enter>", show)
    widget.bind("<Leave>", hide)


LOG_BOX_LINES = 15    # 底部运行日志区可见行数（右侧带滚动条可回看历史）
RESULT_VISIBLE_ROWS = 6   # 测试结果区只显示 6 条数据行（表头另计，超出可滚动）
RESULT_HEADER_H = 40      # 结果表表头最小高度（px）；表头行数更多时按行数自动加高
RESULT_HEADER_LINE_H = 17  # 表头单行文本高度（px），用于按行数估算表头高度
RESULT_ROW_H = 23         # 结果表单行实测高度（px）


def _close_tk_window(win):
    """安全关闭 Tk 窗口（重复/已销毁均忽略），供自动消失提示框定时回调使用。"""
    try:
        if win.winfo_exists():
            win.destroy()
    except Exception:
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CMW500 LTE 射频指标测试工具（信令模式）")
        # 结果表固定 6 行 + 日志 15 行，窗口按内容高度收紧（不再留大片空白）
        self.geometry("1240x940")
        self.minsize(1040, 860)

        self.scpi = None
        self.psu = None
        self.initialized = False
        self.last_result = None
        self.msg_queue = queue.Queue()
        self.busy = False
        self._at_busy = False
        self._at_port_detected = ""
        # AT 串口全局互斥：探测/发送/测量掉线自动恢复共用模块级锁，
        # 避免手动发送与测量恢复同时开同一 COM 口报"资源在使用中"
        self._at_lock = _AT_SER_LOCK

        self.cfg_loaded = load_config()
        self._init_vars()
        try:
            self.run_log = RunLogger()
        except Exception as exc:
            self.run_log = None
            print("运行日志初始化失败:", exc)

        self._build_ui()
        self.after(80, self._poll_queue)
        # 启动时静默探测 AT 口（后台线程，结果经队列回填下拉框）
        self.after(200, lambda: self._detect_at_ports(False))
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------- 变量 ----------------
    def _init_vars(self):
        c = self.cfg_loaded
        def V(key):
            return tk.StringVar(value=str(c.get(key, CONFIG_DEFAULTS.get(key, ""))))
        def B(key):
            return tk.BooleanVar(value=bool(c.get(key, CONFIG_DEFAULTS.get(key, False))))

        # CMW500 行
        self.ip_var = V("cmw_ip")
        self.port_var = V("cmw_port")
        self.rf_var = V("rf")
        self.rx_var = V("rx")

        # 电源行：默认启用，界面不再提供"启用"勾选框（固定 True，不读 config 旧值）
        self.psu_enable_var = tk.BooleanVar(value=True)
        self.psu_host_var = V("psu_ip")
        self.psu_port_var = V("psu_port")
        self.psu_channel_var = V("psu_channel")
        self.psu_voltage_var = V("psu_voltage")
        # psu_readv 不再有勾选框：电压/电流默认记录，下发给 True

        # 测试参数（高级）
        self.freq_var = V("freq")
        self.enp_var = V("enp")
        self.eatt_var = V("eatt")
        self.umar_var = V("umar")
        self.rsrp_var = V("rsrp")
        self.pmax_var = V("pmax")
        self.dl_rb_var = V("dl_rb")
        self.dl_mod_var = V("dl_mod")
        self.sens_start_var = V("sens_start")
        self.sens_stop_var = V("sens_stop")
        self.sens_step_var = V("sens_step")
        self.sens_near_var = V("sens_near")
        self.sens_coarse_var = V("sens_coarse")
        self.sens_hotspot_margin_var = V("sens_hotspot_margin")
        self.sens_hotspot_span_var = V("sens_hotspot_span")
        self.sens_throughput_var = V("sens_throughput")
        self.bler_subframes_var = V("bler_subframes")
        self.bler_timeout_var = V("bler_timeout")
        self.tpc_settle_var = V("tpc_settle")
        self.stats_var = V("stats")

        self.evm_lim_var = V("evm_lim")
        self.aclr_lim_var = V("aclr_lim")
        self.freq_lim_var = V("freq_lim")
        self.sens_lim_var = V("sens_lim")

        self.metric_vars = {}
        for k in METRIC_KEYS:
            self.metric_vars[k] = tk.BooleanVar(value=bool(c["metrics"].get(k, True)))
        # 测试场景（TX/RX 公用）：Pmax / Pmin / Pmax+Pmin
        _mode = str(c.get("power_scenario", c.get("sens_power_mode", "high")) or "high")
        if _mode not in SENS_POWER_MODES:
            _mode = "high"
        self.sens_power_mode_var = tk.StringVar(value=_mode)
        self.sens_low_power_var = V("sens_low_power")
        # 全信道扫描（RX）：勾选后枚举所选频段/带宽下全部合法 DL EARFCN 逐点探底
        self.full_scan_var = tk.BooleanVar(value=bool(c.get("full_scan", False)))

        # AT 控制行：串口选择（空/auto=自动识别）+ 手动 debug 指令输入
        self.at_port_var = V("at_port")
        self.at_cmd_var = V("at_cmd")

        saved_bands = set((c.get("bands") or []))
        saved_band_bws = normalize_band_bws(c)   # 兼容旧版全局 bws
        self.band_vars = {b: tk.BooleanVar(value=(b in saved_bands)) for b in ALL_BANDS}
        # 每个频段独立带宽：{band: {bw: BooleanVar}}；不支持的带宽恒 False
        self.band_bw_vars = {
            b: {w: tk.BooleanVar(value=(w in set(saved_band_bws.get(b) or [])))
                for w in BW_CHOICES}
            for b in ALL_BANDS}
        self.band_bw_summary_var = tk.StringVar(value="")
        # 表头带宽总控：点选某个带宽 → 对所有「已勾选且支持该带宽」的频段一次生效
        self.head_bw_vars = {w: tk.BooleanVar(value=False) for w in BW_CHOICES}
        self.head_bw_cbs = {}

        self.status_var = tk.StringVar(value="就绪")
        self.conn_label = None
        self.psu_label = None

    # ---------------- UI ----------------
    def _build_ui(self):
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
        self._install_checkmark_style(style)

        root = ttk.Frame(self, padding=(10, 2))
        root.pack(fill="both", expand=True)

        # ===== 第1行：CMW500 配置 + 操作按钮（并排两个容器，grid 底部对齐） =====
        top_row = ttk.Frame(root)
        top_row.pack(fill="x", pady=(0, 6))
        top_row.columnconfigure(0, weight=1)  # f1 自然撑满左边
        # CMW500 配置框
        f1 = ttk.LabelFrame(top_row, text="CMW500 配置（网线 / Raw Socket 5025）", padding=(8, 5))
        f1.grid(row=0, column=0, sticky="w")
        ttk.Label(f1, text="IP:").grid(row=0, column=0, sticky="w", padx=(2, 4))
        ttk.Entry(f1, textvariable=self.ip_var, width=15).grid(row=0, column=1, sticky="w")
        ttk.Label(f1, text="端口:").grid(row=0, column=2, sticky="w", padx=(8, 4))
        ttk.Entry(f1, textvariable=self.port_var, width=7).grid(row=0, column=3, sticky="w")
        ttk.Label(f1, text="RF口:").grid(row=0, column=4, sticky="w", padx=(16, 4))
        ttk.Combobox(f1, textvariable=self.rf_var, values=RF_CONNECTORS,
                     width=8, state="readonly").grid(row=0, column=5, sticky="w")
        ttk.Label(f1, text="RX:").grid(row=0, column=6, sticky="w", padx=(8, 4))
        ttk.Combobox(f1, textvariable=self.rx_var, values=RX_CONNECTORS,
                     width=6, state="readonly").grid(row=0, column=7, sticky="w")
        self.conn_label = ttk.Label(f1, text="● 未连接", foreground="#888888")
        self.conn_label.grid(row=0, column=8, sticky="w", padx=(12, 2))

        # AT 控制 Frame（手动 debug 用：AT 口选择/自动识别 + 指令输入 + 发送）。
        # grid sticky="s" 让底部与 f1 内控件对齐。
        f_at = ttk.LabelFrame(top_row, text="AT 控制（USB 串口 / 手动 debug）", padding=(8, 5))
        f_at.grid(row=0, column=1, sticky="s", padx=(10, 0))
        ttk.Label(f_at, text="AT口:").grid(row=0, column=0, sticky="w", padx=(2, 4))
        self.at_port_cb = ttk.Combobox(f_at, textvariable=self.at_port_var,
                                       values=["auto"], width=10)
        self.at_port_cb.grid(row=0, column=1, sticky="w")
        self.btn_at_detect = ttk.Button(f_at, text="检测", width=4,
                                        command=lambda: self._detect_at_ports(True))
        self.btn_at_detect.grid(row=0, column=2, sticky="w", padx=(4, 0))
        ttk.Label(f_at, text="指令:").grid(row=0, column=3, sticky="w", padx=(10, 4))
        self.at_cmd_entry = ttk.Entry(f_at, textvariable=self.at_cmd_var, width=22)
        self.at_cmd_entry.grid(row=0, column=4, sticky="w")
        self.at_cmd_entry.bind("<Return>", lambda _e: self._send_at_cmd())
        self.btn_at_send = ttk.Button(f_at, text="发送", width=5,
                                      command=self._send_at_cmd)
        self.btn_at_send.grid(row=0, column=5, sticky="w", padx=(4, 2))

        # ===== 第2行：电源配置 =====
        f2 = ttk.LabelFrame(root, text="电源配置（Keysight E36312A / LAN 5025）", padding=(8, 5))
        f2.pack(fill="x", pady=(0, 6))
        # 电源默认启用，去掉"启用"勾选框；各控件左对齐、无缝隙
        ttk.Label(f2, text="IP:").grid(row=0, column=0, sticky="w", padx=(2, 4))
        ttk.Entry(f2, textvariable=self.psu_host_var, width=15).grid(row=0, column=1, sticky="w")
        ttk.Label(f2, text="端口:").grid(row=0, column=2, sticky="w", padx=(8, 4))
        ttk.Entry(f2, textvariable=self.psu_port_var, width=7).grid(row=0, column=3, sticky="w")
        ttk.Label(f2, text="通道:").grid(row=0, column=4, sticky="w", padx=(16, 4))
        ch = ttk.Combobox(f2, textvariable=self.psu_channel_var, values=list(PSU_CHANNELS),
                          width=8, state="readonly")
        ch.grid(row=0, column=5, sticky="w")
        ttk.Label(f2, text="电压(V):").grid(row=0, column=6, sticky="w", padx=(8, 4))
        ttk.Entry(f2, textvariable=self.psu_voltage_var, width=9).grid(row=0, column=7, sticky="w")
        # 电压范围提示已去掉（界面更紧凑）；范围校验仍在 _read_cfg 里按 PSU_VOLT_RANGE 执行
        self.psu_label = ttk.Label(f2, text="● 未连接", foreground="#888888")
        self.psu_label.grid(row=0, column=8, sticky="w", padx=(10, 2))

        # 操作按钮 Frame：靠右放置，末按钮右缘与上方 AT 控制区【发送】按钮右缘对齐
        #   f_at 发 送 右缘 = top_row 右缘 -10(f_at padding) -2(btn padx)；
        #   f2 内容右缘 = root 内容右缘 -8(f2 padding)，fbtn sticky="e" 贴右缘，
        #   末按钮 padx 右=2 → 与发送按钮右缘重合。
        fbtn = ttk.Frame(f2)
        fbtn.grid(row=0, column=9, sticky="se", padx=(10, 0))
        f2.columnconfigure(9, weight=1)
        self.btn_init = ttk.Button(fbtn, text="初始化仪表", command=self.initialize, width=10)
        self.btn_init.pack(side="left", padx=(2, 4))
        self.btn_measure = ttk.Button(fbtn, text="开始测量", command=self.start_measurement,
                                      width=10)
        self.btn_measure.pack(side="left", padx=(0, 4))
        self.btn_open_cfg = ttk.Button(fbtn, text="打开配置", command=self._open_config,
                                       width=10)
        self.btn_open_cfg.pack(side="left", padx=(0, 4))
        self.btn_open_specs = ttk.Button(fbtn, text="打开频段表", command=self._open_specs,
                                         width=10)
        self.btn_open_specs.pack(side="left", padx=(0, 2))

        # ===== 中间：频段 → 带宽 → 指标，自上而下依次分布 =====
        mid = ttk.Frame(root)
        mid.pack(fill="x", pady=(2, 6))

        # ① 工作频段与带宽：每个频段一行，只列出该频段支持的带宽
        #    可选带宽 / 低中高信道全部取自本目录 specs\LTE\bands.json
        self._build_band_bw_matrix(mid)

        # ②③ TX / RX 指标放在同一行，各用一个 LabelFrame 包装，节约上下空间
        metrics_row = ttk.Frame(mid)
        metrics_row.pack(fill="x")

        # ① 测试场景（TX/RX 公用）：选 Pmax / Pmin / Pmax+Pmin 后再勾指标，
        #    所有勾选的 TX/RX 指标都在所选场景下测量，结果列带档位标注。
        bf_scenario = ttk.LabelFrame(metrics_row, text="① 测试场景", padding=(8, 5))
        bf_scenario.pack(side="left", fill="y", padx=(0, 8))
        sframe = ttk.Frame(bf_scenario)
        sframe.pack(fill="x")
        for j, mode in enumerate(SENS_POWER_MODES):
            ttk.Radiobutton(sframe, text=SENS_POWER_LABELS[mode], value=mode,
                            variable=self.sens_power_mode_var,
                            command=self._on_sens_mode_change).grid(
                row=0, column=j, sticky="w", padx=(0, 10), pady=3)
        ttk.Label(sframe, text="Pmin:").grid(row=0, column=3, sticky="w", padx=(4, 4))
        self.sens_low_power_entry = ttk.Entry(sframe, width=6,
                                               textvariable=self.sens_low_power_var)
        self.sens_low_power_entry.grid(row=0, column=4, sticky="w", padx=(0, 2))
        ttk.Label(sframe, text="(dBm)").grid(row=0, column=5, sticky="w", padx=(0, 6))
        # Pmin 含义说明：空=最小功率(MINPower)；填值=Closed Loop 功率
        self.pmin_hint = ttk.Label(sframe,
            text="空=MINPower | 填值=闭环功率 (-50~23 dBm)",
            foreground="#4a9eda")
        self.pmin_hint.grid(row=0, column=6, sticky="w", padx=(0, 4))

        # ② TX 指标（按测试场景的功率档测上行）—— 勾选哪些就测哪些
        bf_tx = ttk.LabelFrame(metrics_row, text="② TX 指标", padding=(8, 5))
        bf_tx.pack(side="left", fill="y", padx=(0, 8))
        mframe = ttk.Frame(bf_tx)
        mframe.pack(fill="x")
        for i, k in enumerate(TX_METRIC_KEYS):
            ttk.Checkbutton(mframe, text=METRIC_LABELS[k],
                            variable=self.metric_vars[k],
                            command=self._on_metrics_change).grid(
                row=0, column=i, sticky="w", padx=(2, 14), pady=3)

        # ③ RX 指标
        bf_rx = ttk.LabelFrame(metrics_row, text="③ RX 指标", padding=(8, 5))
        bf_rx.pack(side="left", fill="y")
        rframe = ttk.Frame(bf_rx)
        rframe.pack(fill="x")
        # 全信道扫描放最前（灵敏度探底之前）
        self.full_scan_cb = ttk.Checkbutton(
            rframe, text="全信道扫描", variable=self.full_scan_var)
        self.full_scan_cb.grid(row=0, column=0, sticky="w", padx=(2, 12), pady=3)
        _attach_full_scan_tip(self.full_scan_cb)
        ttk.Checkbutton(rframe, text=METRIC_LABELS["sensitivity"],
                        variable=self.metric_vars["sensitivity"],
                        command=self._on_metrics_change).grid(
            row=0, column=1, sticky="w", padx=(2, 16), pady=3)
        # 外部衰减（旧版手动输入框已移除，改用 freq_loss_segments 按频段/信道
        # 自动插值补偿 EATT:OUTP/EATT:INP，见 retune_cell + else 快速路径）

        # 高级参数不再显示，使用 config.json / 默认值

        # idn_label 仅保留对象供连接逻辑写入（界面不再显示仪器信息，节省空间）
        self.idn_label = ttk.Label(root, text="")

        # ===== 结果区（Frame+Label 网格：每个单元格独立 Label，FAIL 单元格单独标红） =====
        # 用 Frame+Label 代替 Treeview：Treeview 只支持行级 tag 无法实现单元格级标红。
        self.res_frame = ttk.LabelFrame(root, text="测试结果",
                             padding=(6, 2))
        # 不再 expand：结果表高度固定为 6 行，窗口高度随之收紧
        self.res_frame.pack(fill="x")
        # 耗时统计原本单独占一行（空文本时留下红框所示空白），现改为写入状态栏，
        # 结果区不再预留该行。保留对象以兼容 _show_result 的 config 调用。
        self.elapsed_label = ttk.Label(self.res_frame, text="",
                                       foreground="#4a9eda")
        self._build_result_grid(self.cfg_loaded)
        # 初始化功率工况输入框可用状态
        self.after(0, self._on_sens_mode_change)

        # 底部状态栏：先 pack side="bottom"，切最底部空间
        ttk.Label(self, textvariable=self.status_var, anchor="w",
                  relief="sunken", padding=(6, 2)).pack(fill="x", side="bottom")
        # 日志区：固定 15 行高，右侧垂直滚动条可回看历史
        log_frame = ttk.Frame(self)
        log_frame.pack(fill="x", pady=(1, 0))
        self.log_box = tk.Text(log_frame, height=LOG_BOX_LINES, wrap="word",
                               bg="#1e1e1e", fg="#d4d4d4",
                               insertbackground="#d4d4d4")
        log_sb = ttk.Scrollbar(log_frame, orient="vertical",
                               command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=log_sb.set)
        self.log_box.pack(side="left", fill="both", expand=True)
        log_sb.pack(side="right", fill="y")

    # ---------------- 辅助 UI ----------------
    def _install_checkmark_style(self, style):
        """clam 主题的勾选框是 ×，用勾形图片替换 indicator，显示为 ✓。

        PhotoImage 必须保持引用，否则会被 GC 导致图标空白，故挂到 self。
        """
        def make_img(checked):
            im = tk.PhotoImage(width=14, height=14)
            im.put("#ffffff", (0, 0, 13, 13))
            for i in range(14):
                im.put("#8a8a8a", (i, 0)); im.put("#8a8a8a", (i, 13))
                im.put("#8a8a8a", (0, i)); im.put("#8a8a8a", (13, i))
            if checked:
                for x, y in [(3, 7), (3, 8), (4, 8), (4, 9), (5, 9), (5, 10),
                             (6, 9), (6, 8), (7, 7), (7, 6), (8, 5), (8, 4),
                             (9, 3), (9, 4), (10, 3), (10, 2)]:
                    im.put("#1a7d32", (x, y))
            return im

        try:
            self._chk_off = make_img(False)
            self._chk_on = make_img(True)
            style.element_create(
                "Checkmark.indicator", "image",
                self._chk_off, ("selected", self._chk_on),
                ("disabled", self._chk_off), ("selected disabled", self._chk_on))
            style.layout("TCheckbutton", [
                ("Checkbutton.padding", {"sticky": "nswe", "children": [
                    ("Checkmark.indicator", {"side": "left", "sticky": ""}),
                    ("Checkbutton.focus", {"side": "left", "sticky": "w", "children": [
                        ("Checkbutton.label", {"sticky": "nswe"})]})]})])
        except Exception as exc:
            # 样式定制失败不影响功能（退化为 clam 原生勾选框）
            # 此处 log_box 尚未创建，只写文件日志/控制台，不走 self.log()
            msg = "[样式] 勾形勾选框设置失败，使用默认: %s" % exc
            if getattr(self, "run_log", None):
                self.run_log.write(msg)
            print(msg)

    def _scroll_checklist(self, parent, options, var_dict, ncols=6, height=92):
        outer = ttk.Frame(parent)
        canvas = tk.Canvas(outer, height=height, highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=inner, anchor="nw")
        for i, opt in enumerate(options):
            ttk.Checkbutton(inner, text=opt, variable=var_dict[opt], width=7).grid(
                row=i // ncols, column=i % ncols, sticky="w", padx=2, pady=1)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        outer.columnconfigure(0, weight=1)
        return outer

    # ---------------- 频段 × 各自带宽 矩阵 ----------------
    def _build_band_bw_matrix(self, parent):
        """每个频段一行，单独勾选该频段要测的带宽。

        可选带宽完全取自本目录 data\\LTE\\bands.json（band_supported_bws）；
        不支持的带宽显示灰色 — 且不创建勾选框，从源头杜绝非法组合。
        """
        ncol = len(BW_CHOICES)
        box = ttk.LabelFrame(parent, text="", padding=(8, 5))
        box.pack(fill="x", pady=(0, 5))

        # 标题与两个批量按钮同一行：标题文字放入 Frame，按钮紧随其后
        bar = ttk.Frame(box)
        bar.pack(fill="x", pady=(0, 4))
        ttk.Label(bar, text="① 工作频段与带宽",
                  font=("TkDefaultFont", 10, "bold")).pack(side="left", padx=(2, 12))
        ttk.Button(bar, text="全选带宽", width=8,
                   command=lambda: self._bulk_band_bws(True)).pack(
                       side="left", padx=(0, 6))
        ttk.Button(bar, text="清空", width=8,
                   command=self._clear_all_selection).pack(
                       side="left", padx=(0, 6))
        ttk.Label(bar, textvariable=self.band_bw_summary_var,
                  foreground="#2e7d32", anchor="e").pack(side="right")

        hint = ttk.Frame(box)
        hint.pack(fill="x", pady=(0, 4))
        ttk.Label(hint,
                  text="批量：先勾选多个频段，再点表头带宽复选框 → 一次性全选/取消该带宽；"
                       "勾选频段默认只勾 10MHz（不支持则取最接近的）；"
                       "「清空全部选择」会连已勾选频段一起取消；"
                       "灰色表头 = 当前无可作用频段。",
                  foreground="#4a9eda").pack(side="left")

        def setup_cols(fr):
            fr.columnconfigure(0, weight=0, minsize=112)
            for j in range(1, ncol + 1):
                fr.columnconfigure(j, weight=1, minsize=104)

        # 表头固定在滚动区外，滚动时始终可见
        head = ttk.Frame(box)
        head.pack(fill="x", padx=(0, 18))   # 右侧留滚动条宽度，与数据列对齐
        setup_cols(head)
        ttk.Label(head, text="频段", anchor="w").grid(
            row=0, column=0, sticky="w", padx=(2, 4))
        for j, w in enumerate(BW_CHOICES, 1):
            cb = ttk.Checkbutton(head, text=bw_label(w),
                                 variable=self.head_bw_vars[w],
                                 command=lambda ww=w: self._on_head_bw_toggle(ww))
            cb.grid(row=0, column=j, sticky="w", padx=2)
            self.head_bw_cbs[w] = cb

        outer = ttk.Frame(box)
        outer.pack(fill="x")
        # 默认仅可见约 5 行频段（原 9 行占界面太高），其余靠右侧滚动条上下滑动
        canvas = tk.Canvas(outer, height=130, highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        win = canvas.create_window((0, 0), window=inner, anchor="nw")
        setup_cols(inner)
        canvas.configure(yscrollcommand=sb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        outer.columnconfigure(0, weight=1)

        def _sync_width(_e=None):
            try:
                canvas.itemconfig(win, width=canvas.winfo_width())
            except tk.TclError:
                pass
        canvas.bind("<Configure>", _sync_width)
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        def _wheel(e):
            try:
                canvas.yview_scroll(int(-e.delta / 120) * 3, "units")
            except (AttributeError, tk.TclError):
                pass
        def _bind_wheel(_e):
            canvas.bind_all("<MouseWheel>", _wheel)
        def _unbind_wheel(_e):
            canvas.unbind_all("<MouseWheel>")
        for wgt in (canvas, inner):
            wgt.bind("<Enter>", _bind_wheel)
            wgt.bind("<Leave>", _unbind_wheel)

        self.band_row_cbs = {}
        for i, b in enumerate(ALL_BANDS, 1):
            cb = ttk.Checkbutton(
                inner, text=band_label(b), variable=self.band_vars[b], width=12,
                command=lambda bb=b: self._on_band_toggle(bb))
            cb.grid(row=i, column=0, sticky="w", padx=(2, 4), pady=1)
            _attach_band_tip(cb, b)
            self.band_row_cbs[b] = cb
            supported = band_supported_bws(b)
            for j, w in enumerate(BW_CHOICES, 1):
                if w in supported:
                    ttk.Checkbutton(inner, text="", variable=self.band_bw_vars[b][w],
                                    command=self._refresh_band_bw_summary).grid(
                        row=i, column=j, sticky="w", padx=2, pady=1)
                else:
                    ttk.Label(inner, text="—", foreground="#aaaaaa").grid(
                        row=i, column=j, sticky="w", padx=2, pady=1)
        self._refresh_band_bw_summary()

    def _head_bw_targets(self, bw):
        """表头总控作用对象：已勾选且支持该带宽的频段。"""
        return [b for b in ALL_BANDS
                if self.band_vars[b].get() and bw in band_supported_bws(b)]

    def _on_head_bw_toggle(self, bw):
        """点表头某个带宽复选框 → 所有已勾选且支持它的频段一次全选/全取消。"""
        on = self.head_bw_vars[bw].get()
        targets = self._head_bw_targets(bw)
        for b in targets:
            self.band_bw_vars[b][bw].set(on)
        if targets:
            self._set_status("表头批量：%s 已%s %d 个已勾选频段" %
                             (bw_label(bw), "勾选" if on else "取消", len(targets)))
        self._refresh_band_bw_summary()

    def _on_band_toggle(self, band):
        """勾选频段时若该频段一个带宽都没选，默认只勾 10MHz；取消勾选时清空其带宽。

        不支持 10MHz 的频段取最接近的支持带宽（5MHz/3MHz）并在状态栏提示；
        用户已手动勾过带宽的频段保持原样，不覆盖其选择。
        取消频段勾选时一并清掉该频段已勾带宽，避免留下"带宽勾着但频段没选"
        的残留状态（勾回该频段时会再默认补 10MHz）。
        """
        if self.band_vars[band].get():
            if not any(v.get() for v in self.band_bw_vars[band].values()):
                w = default_bw_for_band(band)
                if w:
                    self.band_bw_vars[band][w].set(True)
                    if w != DEFAULT_BW:
                        self._set_status("%s 不支持 10MHz，默认勾选 %s" %
                                         (band_label(band), bw_label(w)))
                    else:
                        self._set_status("%s 默认勾选 %s" %
                                         (band_label(band), bw_label(w)))
        else:
            # 取消频段 → 清空该频段全部带宽勾选（保留残留会造成汇总/批量误判）
            for w in BW_CHOICES:
                self.band_bw_vars[band][w].set(False)
        self._refresh_band_bw_summary()

    def _bulk_band_bws(self, on):
        """批量：已选频段全选各自支持的带宽 / 仅清空所有频段的带宽。

        界面上的「清空全部选择」按钮不走这里，而是调用 _clear_all_selection，
        它在清空带宽的同时还会取消已勾选的频段。
        """
        for b in ALL_BANDS:
            if on and not self.band_vars[b].get():
                continue
            supported = band_supported_bws(b) if on else []
            for w in BW_CHOICES:
                self.band_bw_vars[b][w].set(w in supported)
        self._refresh_band_bw_summary()

    def _clear_all_selection(self):
        """清空全部选择：取消所有已勾选的频段，并清空所有频段的带宽勾选。

        表头批量复选框状态由 _refresh_band_bw_summary 统一同步
        （没有可作用频段时自动置灰并复位）。直接 set 变量不会触发
        Checkbutton 的 command，因此不会误触发 _on_band_toggle 的默认勾选。
        """
        n_bands = sum(1 for b in ALL_BANDS if self.band_vars[b].get())
        for b in ALL_BANDS:
            self.band_vars[b].set(False)
            for w in BW_CHOICES:
                self.band_bw_vars[b][w].set(False)
        self._refresh_band_bw_summary()
        self._set_status("已清空全部选择：取消 %d 个频段及全部带宽勾选" % n_bands)

    def _refresh_band_bw_summary(self):
        bands = [b for b in ALL_BANDS if self.band_vars[b].get()]
        bb = {b: [w for w in BW_CHOICES
                  if w in band_supported_bws(b) and self.band_bw_vars[b][w].get()]
              for b in ALL_BANDS}
        for b, cb in (getattr(self, "band_row_cbs", None) or {}).items():
            cb.configure(text="%s (%d/%d)" %
                         (band_label(b), len(bb[b]), len(band_supported_bws(b))))
        pairs = [(b, w) for b in bands for w in bb[b]]
        points = sum(len(band_channels(b, w)) for b, w in pairs)
        detail = band_bws_summary(bands, bb, sep="；", limit=4)
        self.band_bw_summary_var.set(
            "已选 %d 频段 / %d 组合 / %d 测试点   ｜   %s" %
            (len(bands), len(pairs), points, detail))
        # 表头总控状态同步：全勾=True / 部分或全空=False；无可作用频段则置灰
        for w in BW_CHOICES:
            tgt = self._head_bw_targets(w)
            self.head_bw_vars[w].set(
                bool(tgt) and all(self.band_bw_vars[b][w].get() for b in tgt))
            cb = (getattr(self, "head_bw_cbs", None) or {}).get(w)
            if cb is not None:
                cb.configure(state=("normal" if tgt else "disabled"))

    def _on_sens_mode_change(self):
        """功率工况切换：仅 low/both 时小功率输入框可用；同步刷新结果表首行。"""
        if not hasattr(self, "sens_low_power_entry"):
            return
        mode = self.sens_power_mode_var.get()
        state = "normal" if mode in ("low", "both") else "disabled"
        try:
            self.sens_low_power_entry.config(state=state)
        except Exception:
            pass
        self._on_metrics_change()

    def _on_metrics_change(self):
        """勾选 TX/RX 指标后立刻刷新结果表首行，无需点击开始测试。

        测量进行中（busy）不重建，避免把已插入的数据行清掉。
        """
        if getattr(self, "busy", False):
            return
        if not hasattr(self, "metric_vars") or not hasattr(self, "_res_canvas"):
            return
        try:
            cfg = {
                "metrics": {k: bool(v.get()) for k, v in self.metric_vars.items()},
                "power_scenario": self.sens_power_mode_var.get().strip(),
                "sens_power_mode": self.sens_power_mode_var.get().strip(),
                "psu_enable": True,   # 界面无电源勾选框，固定启用
            }
            self._rebuild_result_columns(cfg)
        except Exception:
            pass

    # ---------------- 日志/状态 ----------------
    def log(self, text):
        if self.run_log:
            self.run_log.write(text)
        if threading.current_thread() is threading.main_thread():
            self._log_direct(text)
        else:
            self.msg_queue.put(("log", text))

    # ---------------- AT 控制（USB 串口，手动 debug） ----------------
    AT_BAUD = 115200          # ASR 平台 USB CDC 虚拟串口，波特率不敏感
    AT_PROBE_HINTS = ("modem", "asr", "at ", "at-", "serial")  # 描述关键词 → 优先探测

    def _detect_at_ports(self, manual=False):
        """枚举 COM 口并逐个做 AT 握手，找出 AT 口；后台线程，结果经队列回 GUI。

        manual=True 时由【检测】按钮触发；启动时静默跑一次（manual=False）。
        """
        if serial is None:
            self.log("[AT] 缺少 pyserial 模块，AT 控制不可用（pip install pyserial）")
            return
        if not self._at_lock.acquire(blocking=False):
            self.log("[AT] 检测跳过：已有 AT 操作（探测/发送）进行中")
            return

        def worker():
            names = []
            found = ""
            try:
                ports = list(serial.tools.list_ports.comports())
                if not ports:
                    self.log("[AT] 未发现任何 COM 口（请检查 DUT USB 是否插好）")
                    return
                # 描述含 modem/asr 等关键词的排前面，探测更快命中
                def prio(p):
                    desc = (p.description or "").lower()
                    return (0 if any(k in desc for k in self.AT_PROBE_HINTS) else 1,
                            p.device)

                def probe_one(device):
                    """单端口 AT 握手；子线程跑、join(6s) 看门狗，卡死则跳过。"""
                    res = {}
                    def run():
                        ser = None
                        try:
                            ser = serial.Serial(device, self.AT_BAUD, timeout=0.8)
                        except Exception as e:
                            res["err"] = str(e)
                            return
                        try:
                            ser.reset_input_buffer()
                            ser.write(b"AT\r\n")
                            t0 = time.time()
                            got = ""
                            while time.time() - t0 < 2.0:
                                data = ser.readline()
                                if not data:
                                    continue
                                got += data.decode(errors="replace")
                                if "OK" in got or "ERROR" in got:
                                    break
                            res["ok"] = "OK" in got
                        except Exception:
                            res["err"] = "io"
                        finally:
                            try:
                                ser.close()
                            except Exception:
                                pass
                    t = threading.Thread(target=run, daemon=True)
                    t.start()
                    t.join(6.0)
                    if t.is_alive():
                        self.log("[AT] %s 探测超时（口被占用或模组卡死），跳过" % device)
                    elif "err" in res:
                        self.log("[AT] %s 打不开（被其它串口工具占用？）： %s"
                                 % (device, res["err"].splitlines()[0] if res["err"] else ""))
                    return bool(res.get("ok"))

                for p in sorted(ports, key=prio):
                    names.append(p.device)
                    if found:
                        continue   # 已找到 AT 口，其余端口只列出不再探测
                    try:
                        if probe_one(p.device):
                            found = p.device
                            self.log("[AT] %s 是 AT 口（%s），已选中"
                                     % (p.device, (p.description or "").strip()))
                    except Exception:
                        pass
                if not found:
                    self.log("[AT] 未探测到响应 AT 的串口（非 AT 口通常无响应；"
                             "若模组刚被 CFUN=0 卡死，请给 DUT 断电重启后再试）")
            except Exception as e:
                self.log("[AT] 端口检测异常: %s" % type(e).__name__)
            finally:
                self._at_lock.release()
                self.msg_queue.put(("at_ports", (names, found)))
        threading.Thread(target=worker, daemon=True).start()

    def _send_at_cmd(self):
        """发送输入框里的 AT 指令（后台线程执行，日志实时回显）。

        支持多条指令依次执行：用 ; 分隔（如 AT+CFUN=0; AT+CFUN=1），
        同一串口会话内按顺序逐条下发，单条失败不中断后续指令。
        """
        if getattr(self, "_at_busy", False):
            return
        raw = self.at_cmd_var.get().strip()
        if not raw:
            return
        cmds = [c.strip() for c in raw.replace("；", ";").split(";") if c.strip()]
        if not cmds:
            return
        if serial is None:
            self.log("[AT] 缺少 pyserial 模块，AT 控制不可用（pip install pyserial）")
            return
        port = self.at_port_var.get().strip()
        if not port or port.lower() == "auto":
            port = getattr(self, "_at_port_detected", "")
            if not port:
                self.log("[AT] 未确定 AT 口：请先点【检测】，或在 AT口 框手动填 COMx")
                return
            self.at_port_var.set(port)
        self._at_busy = True
        self.btn_at_send.config(state="disabled")
        if len(cmds) > 1:
            self.log("[AT] 批量执行 %d 条指令（依次下发）" % len(cmds))
        threading.Thread(target=self._at_worker, args=(port, cmds), daemon=True).start()

    def _at_worker(self, port, cmds):
        """在同一串口会话内依次下发多条指令；每条读到 OK/ERROR 或 8s 超时。

        - 单条失败（异常/超时）不中断后续指令；
        - 全程持 _at_lock（与探测互斥）；打开失败自动重试 3 次；
        - 外层 join(15s×条数) 看门狗：串口/模组卡死时释放按钮，
          残留阻塞线程为 daemon，随进程退出；
        - 结束后经队列在主线程保存输入框内容（at_cmd 持久化）。
        """
        if not self._at_lock.acquire(timeout=15):
            self.log("[AT] 发送放弃：探测/另一条指令占用串口超过 15s")
            self._at_busy = False
            self.msg_queue.put(("at_done", None))
            return
        try:
            def run():
                ser = None
                try:
                    # 打开重试：口被短暂占用（检测残留句柄等）时等 0.8s 再试
                    for attempt in range(3):
                        try:
                            ser = serial.Serial(port, self.AT_BAUD, timeout=1.0)
                            break
                        except serial.SerialException as oe:
                            if attempt == 2:
                                self.log("[AT] %s 打不开（口被占用）： %s\n"
                                         "[AT] 建议：先关闭占用该口的其它串口工具\n"
                                         "（如单独 AT 工具/串口助手），Windows 串口独占、"
                                         "两程序不能同开一个口" % (port, oe))
                                return
                            time.sleep(0.8)
                except Exception as e:
                    self.log("[AT] %s 打开失败: %s %s" % (port, type(e).__name__, e))
                    return
                try:
                    time.sleep(0.3)
                    ser.reset_input_buffer()
                    for i, cmd in enumerate(cmds):
                        tag = ("[%d/%d] " % (i + 1, len(cmds))) if len(cmds) > 1 else ""
                        self.log("[AT] %s%s > %s" % (tag, port, cmd))
                        try:
                            ser.write((cmd + "\r\n").encode())
                        except Exception as e:
                            # 写失败（模组重启掉口等）：本条失败，继续下一条
                            self.log("[AT] %s < 写入失败: %s" % (port, e))
                            continue
                        t0 = time.time()
                        while time.time() - t0 < 8.0:
                            try:
                                data = ser.readline()
                            except Exception as e:
                                self.log("[AT] %s < 读取失败: %s" % (port, e))
                                break
                            if not data:
                                continue
                            line = data.decode(errors="replace").strip()
                            if line:
                                self.log("[AT] %s < %s" % (port, line))
                                if line in ("OK", "ERROR") or line.startswith("+CME ERROR"):
                                    break
                        if len(cmds) > 1:
                            time.sleep(0.3)   # 条与条之间留模组处理间隙
                finally:
                    try:
                        ser.close()
                    except Exception:
                        pass
            t = threading.Thread(target=run, daemon=True)
            t.start()
            t.join(max(15.0, 15.0 * len(cmds)))
            if t.is_alive():
                self.log("[AT] %s 卡死（串口/模组无响应），已跳过本次发送" % port)
        finally:
            self._at_lock.release()
            self._at_busy = False
            self.msg_queue.put(("at_done", None))
            self.msg_queue.put(("at_save", None))


    def _log_direct(self, text):
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_box.insert("end", "[%s] %s\n" % (ts, text))
        self.log_box.see("end")

    def _set_status(self, text):
        self.status_var.set(text)

    def _set_busy(self, busy):
        state = "disabled" if busy else "normal"
        for b in (self.btn_init, self.btn_measure):
            b.config(state=state)

    # ---------------- 配置读取 ----------------
    def _num(self, name, var, cast=float):
        raw = var.get().strip()
        if not raw:
            raise ValueError("%s 不能为空" % name)
        try:
            return cast(raw)
        except ValueError:
            raise ValueError("%s 数值无效: %s" % (name, raw))

    def _tdd_capability(self):
        """缓存式探测仪表 TDD 授权（仅探测一次）。返回 (capable, detail)。"""
        if getattr(self, "_tdd_cap_cache", None) is not None:
            return self._tdd_cap_cache
        try:
            # 必须读用户配置（self.cfg_loaded），只读 CONFIG_DEFAULTS 会让
            # config.json 里的 tdd_detect=0 失效，仍误探测并剔除 TDD 组合
            _detect = (self.cfg_loaded or {}).get(
                "tdd_detect", CONFIG_DEFAULTS.get("tdd_detect", "0"))
            if str(_detect) == "0":
                self._tdd_cap_cache = (True, "TDD 探测已禁用（按仪表支持处理）")
                return self._tdd_cap_cache
        except Exception:
            pass
        scpi = None
        try:
            # 独立临时连接：不触碰 self.scpi / 连接指示灯（探测后立即关闭）
            host = self.ip_var.get().strip()
            port = int(self.port_var.get().strip() or 5025)
            scpi = ScpiTransport(host, port, timeout=10.0, eol="\n")
            scpi.connect()
            capable, detail = detect_tdd_capability(scpi, self.log)
        except Exception as e:
            capable, detail = False, "TDD 探测失败: %s" % e
        finally:
            if scpi is not None:
                try:
                    scpi.close()
                except Exception:
                    pass
        self._tdd_cap_cache = (capable, detail)
        self.log("[TDD] 授权探测：%s（%s）" %
                 ("支持" if capable else "不支持", detail))
        return self._tdd_cap_cache

    def _read_cfg(self):
        cfg = dict(CONFIG_DEFAULTS)
        cfg.update({
            "cmw_ip": self.ip_var.get().strip(),
            "cmw_port": self.port_var.get().strip(),
            "rf": self.rf_var.get().strip(),
            "rx": self.rx_var.get().strip(),
            "psu_enable": True,   # 电源默认启用，界面无勾选框
            "psu_ip": self.psu_host_var.get().strip(),
            "psu_port": self.psu_port_var.get().strip(),
            "psu_channel": self.psu_channel_var.get().strip().upper(),
            "psu_voltage": self.psu_voltage_var.get().strip(),
            "psu_readv": True,   # 电压/电流默认记录，无勾选框
            "psu_samples": PSU_DEFAULT_SAMPLES,
            "freq": self.freq_var.get().strip(),
            "enp": self.enp_var.get().strip(),
            "eatt": self.eatt_var.get().strip(),
            "umar": self.umar_var.get().strip(),
            "rsrp": self.rsrp_var.get().strip(),
            "pmax": self.pmax_var.get().strip(),
            "dl_rb": self.dl_rb_var.get().strip(),
            "dl_mod": self.dl_mod_var.get().strip().upper(),
            "sens_start": self.sens_start_var.get().strip(),
            "sens_stop": self.sens_stop_var.get().strip(),
            "sens_step": self.sens_step_var.get().strip(),
            "sens_near": self.sens_near_var.get().strip(),
            "sens_coarse": self.sens_coarse_var.get().strip(),
            "sens_hotspot_margin": self.sens_hotspot_margin_var.get().strip(),
            "sens_hotspot_span": self.sens_hotspot_span_var.get().strip(),
            "sens_throughput": self.sens_throughput_var.get().strip(),
            "bler_subframes": self.bler_subframes_var.get().strip(),
            "bler_timeout": self.bler_timeout_var.get().strip(),
            "tpc_settle": self.tpc_settle_var.get().strip(),
            "stats": self.stats_var.get().strip(),
            "evm_lim": self.evm_lim_var.get().strip(),
            "aclr_lim": self.aclr_lim_var.get().strip(),
            "freq_lim": self.freq_lim_var.get().strip(),
            "sens_lim": self.sens_lim_var.get().strip(),
            "power_scenario": self.sens_power_mode_var.get().strip(),
            "sens_power_mode": self.sens_power_mode_var.get().strip(),  # 兼容旧调用
            "sens_low_power": self.sens_low_power_var.get().strip(),
            "full_scan": bool(self.full_scan_var.get()),
            "at_port": self.at_port_var.get().strip(),
            "at_cmd": self.at_cmd_var.get().strip(),
            # GUI 启动时 AT 自动识别结果（供测量线程的掉线自动恢复使用）
            "at_port_detected": getattr(self, "_at_port_detected", ""),
            "metrics": {k: bool(v.get()) for k, v in self.metric_vars.items()},
        })
        cfg["bands"] = [b for b in ALL_BANDS if self.band_vars[b].get()]
        # 每个频段各自勾选的带宽（只保留该频段支持的，杜绝非法信道）
        band_bws = {}
        for b in ALL_BANDS:
            sup = band_supported_bws(b)
            band_bws[b] = [w for w in BW_CHOICES
                           if w in sup and self.band_bw_vars[b][w].get()]
        cfg["band_bws"] = band_bws
        # bws 保留为并集（报告/旧字段兼容），不再是组合枚举依据
        cfg["bws"] = [w for w in BW_CHOICES
                      if any(w in band_bws.get(b, []) for b in cfg["bands"])]
        # 测试项 = 每个已选频段 × 该频段自己的带宽 × 信道
        # 默认低/中/高三信道；勾选「全信道扫描」时枚举全部合法 DL EARFCN。
        # 信道按 (频段, 带宽) 从本目录 specs\LTE\bands.json 的 R19 数据取
        combos, skipped = build_combos(cfg["bands"], band_bws,
                                       full_scan=bool(cfg.get("full_scan")))
        # TDD 授权过滤：仪表无 KS550 时 TDD band/CHAN 写入会被 -221/-222 拒绝，
        # 但 UE 仍在网 → 会产出"看着正常实则跑旧 FDD 小区"的无效数据（112703.log
        # 55 组）。探测一次（缓存），无授权则剔除全部 TDD 组合并明确告知。
        if any(is_tdd_band(c[0]) for c in combos):
            capable, detail = self._tdd_capability()
            if not capable:
                tdd_bands = sorted({c[0] for c in combos if is_tdd_band(c[0])})
                combos = [c for c in combos if not is_tdd_band(c[0])]
                skipped = skipped + [combo_display(b, w) for b in tdd_bands
                                     for w in band_bws.get(b, [])]
                self.log("[TDD] 跳过 TDD 频段 %s：%s（不会产出无效数据）" %
                         (",".join(band_label(b) for b in tdd_bands), detail))
        cfg["combos"] = combos
        cfg["skipped_combos"] = skipped

        if not cfg["bands"]:
            raise ValueError("请至少勾选一个频段")
        if not cfg["bws"]:
            raise ValueError(
                "请为勾选的频段选择至少一个带宽。\n"
                "每个频段可单独勾选不同带宽；只有该频段支持的带宽可勾选"
                "（灰色 — 表示不支持）。")
        if not combos:
            raise ValueError(
                "所选频段与带宽没有有效组合。\n"
                "被跳过: %s\n请改选该频段支持的带宽（各频段支持带宽见"
                "《LTE频段_带宽_高中低信道对照表.txt》）。" % (", ".join(skipped) or "无"))
        if not any(cfg["metrics"].values()):
            raise ValueError("请至少勾选一个测试指标")
        # 测试场景（TX/RX 公用）校验 + 新旧键同步
        _scenario = cfg.get("power_scenario", cfg.get("sens_power_mode", "high"))
        if _scenario not in SENS_POWER_MODES:
            raise ValueError("测试场景无效: %s" % _scenario)
        cfg["power_scenario"] = cfg["sens_power_mode"] = _scenario
        if cfg["sens_power_mode"] in ("low", "both"):
            raw = (cfg.get("sens_low_power") or "").strip()
            if raw:  # 非空才校验范围；空 = MINPower 不需验证
                try:
                    lp = float(raw)
                except ValueError:
                    raise ValueError("小功率值无效: %s" % raw)
                if lp < -50 or lp > 23:
                    raise ValueError("小功率值需在 -50 ~ 23 dBm 之间（CMW TPC CLTPower 范围）")
        # 灵敏度判定门限：Throughput 百分比 0~100（默认 95，PASS 条件 TPut ≥ 门限）
        try:
            tp = float(cfg.get("sens_throughput", "95"))
        except (TypeError, ValueError):
            raise ValueError("Throughput 门限无效: %s" % cfg.get("sens_throughput"))
        if tp <= 0 or tp > 100:
            raise ValueError("Throughput 门限需在 0 ~ 100% 之间（默认 95）")
        # 外接衰减器补偿校验（0 = 不接衰减器；当前不接，保持 0）
        try:
            eatt = float(cfg.get("eatt", "0") or 0)
        except (TypeError, ValueError):
            raise ValueError("外部衰减值无效: %s" % cfg.get("eatt"))
        if eatt < 0 or eatt > 80:
            raise ValueError("外部衰减值需在 0 ~ 80 dB 之间")
        if cfg["psu_enable"]:
            if not cfg["psu_ip"]:
                raise ValueError("请填写电源 IP（电源默认启用）")
            if cfg["psu_channel"] not in PSU_CHANNELS:
                raise ValueError("电源通道只能是 P6V/P25V/N25V")
            if cfg["psu_voltage"]:
                lo, hi, hint = PSU_VOLT_RANGE[cfg["psu_channel"]]
                v = float(cfg["psu_voltage"])
                if v < lo or v > hi:
                    raise ValueError("电压 %.3f 超出 %s 范围 %s" % (v, cfg["psu_channel"], hint))
        for key, attr in _VAR_MAP.items():
            self._num(key, getattr(self, attr))
        return cfg

    def save_settings(self):
        try:
            cfg = self._read_cfg()
        except Exception:
            cfg = dict(self.cfg_loaded)
        if save_config(cfg):
            self.cfg_loaded = cfg

    # ---------------- 连接 ----------------
    def _connect_scpi(self):
        host = self.ip_var.get().strip()
        if not host:
            raise ValueError("请填写 CMW500 的 IP 地址")
        port = int(self.port_var.get().strip() or 5025)
        eol = "\n"  # CMW500 及绝大多数现代 SCPI 仪器默认 LF，无需 CRLF 可配置
        scpi = ScpiTransport(host, port, timeout=10.0, eol=eol)
        idn = scpi.connect()
        self.scpi = scpi
        self.conn_label.config(text="● 已连接", foreground="#2e7d32")
        self.idn_label.config(text=idn)
        self.log("[连接] %s:%d %s" % (host, port, idn))
        return scpi

    def _connect_psu(self):
        # 电源默认启用（界面无勾选框），总是连接
        host = self.psu_host_var.get().strip()
        if not host:
            self.log("[电源连接] 跳过：未填写 IP")
            self.psu_label.config(text="● 未连接（跳过）", foreground="#888")
            return None
        port = int(self.psu_port_var.get().strip() or 5025)
        psu = ScpiTransport(host, port, timeout=10.0)
        try:
            idn = psu.connect()
            psu.write("SYST:REM")
        except Exception as e:
            psu.close()
            self.log("[电源连接] ⚠️ 失败：%s:%d — %s" % (host, port, type(e).__name__))
            self.log("[电源连接] 提示：电源功耗测量不可用，但射频指标（灵敏度/EVM/ACLR）不受影响")
            self.psu_label.config(text="⚠️ 连接失败（功耗跳过）", foreground="#e67e22")
            return None
        if not (idn or "").strip():
            # socket 建立但设备没回 IDN：真机 180539.log Run2 曾静默继续，
            # 之后 12 条功耗查询全空才提示"运行中断开"，误导排障方向。
            psu.close()
            self.log("[电源连接] ⚠️ %s:%d 已连接但未返回 IDN（设备无响应/地址错误），"
                     "功耗数据跳过" % (host, port))
            self.log("[电源连接] 提示：请确认电源 IP 与网络可达；射频指标不受影响")
            self.psu_label.config(text="⚠️ 无响应（功耗跳过）", foreground="#e67e22")
            return None
        self.psu = psu
        self.psu_label.config(text="● 已连接", foreground="#2e7d32")
        self.log("[电源连接] %s:%d %s" % (host, port, idn))
        return psu

    def disconnect(self):
        if self.scpi:
            self.scpi.close()
            self.scpi = None
        if self.psu:
            self.psu.close()
            self.psu = None
        self.initialized = False
        self.conn_label.config(text="● 未连接", foreground="#888888")
        self.psu_label.config(text="● 未连接", foreground="#888888")

    # ---------------- ① 初始化仪表 ----------------
    def initialize(self):
        if self.busy:
            messagebox.showinfo("忙", "请等待当前操作结束。")
            return
        try:
            cfg = self._read_cfg()
        except ValueError as exc:
            messagebox.showwarning("输入错误", str(exc))
            return
        self.save_settings()
        self.busy = True
        self._set_busy(True)
        self._set_status("初始化中...")
        threading.Thread(target=self._init_worker, args=(cfg,), daemon=True).start()

    def _init_worker(self, cfg):
        try:
            # 初始化小区必须用第一个测试组合的频段/带宽：
            # 现在每个频段带宽不同，用全局 bws[0] 可能与该频段不匹配。
            first = (cfg.get("combos") or [("", "", "", -1)])[0]
            band, bw = first[0], first[1]
            # 记录小区实际建立的频段/带宽：测量循环用它判断"是否换带宽"，
            # 不能假定它等于开始测量时的首组合——用户可能在初始化后
            # 修改了带宽勾选（真机 180539.log：init 按 10MHz 建小区，
            # 测量首组合却是 1.4MHz，快速路径未重配 RMC → MEV 全 INV）。
            self.cell_actual = (band, bw)
            scpi = self._connect_scpi()
            psu = self._connect_psu()

            cmds, meta = build_signaling_init(cfg, band, bw)
            for cmd in cmds:
                scpi.write(cmd)
                self.log("[SCPI] " + cmd)

            scpi.write(meta["cell_on"])
            self.log("[SCPI] " + meta["cell_on"])
            ok, st = wait_for(scpi, meta["cell_state_q"],
                              cell_ready, 15.0, 0.4, self.log, expect=RE_CELL)
            if not ok:
                raise ScpiError("小区未就绪（需 ON,ADJ）: %s" % st)

            scpi.write(meta["conn_cmd"])
            self.log("[SCPI] " + meta["conn_cmd"])
            ok, st = wait_for(scpi, meta["conn_state_q"],
                              conn_ready, 30.0, 0.5, self.log, expect=RE_PSW)
            self.log("[连接状态] %s（UE 已注册: ATT/ATTached 或 CEST/CESTablished）" % st)
            if not ok:
                # ON 仅表示信号开，绝非注册成功；未注册必须中止，否则测出全 INV 还误报"测完"
                raise ScpiError(
                    "UE 未注册到小区（PSW:STAT=%s）。测量已中止。\n"
                    "可能原因：DUT 未开机/未进入测试模式、DL 信号功率不足、"
                    "频段与 UE 不匹配、鉴权/安全设置不一致。" % st)

            if psu:
                for cmd in build_psu_voltage_commands(cfg):
                    psu.write(cmd)
                    self.log("[PSU] " + cmd)

            self.msg_queue.put(("init_done", "仪表初始化完成（小区 ON,ADJ，UE 已注册）"))
        except Exception:
            self.msg_queue.put(("error", traceback.format_exc()))

    # ---------------- ② 开始测量 ----------------
    def start_measurement(self):
        if self.busy:
            messagebox.showinfo("忙", "上一次测量尚未结束。")
            return
        if not self.initialized:
            messagebox.showwarning("未初始化", "请先点上方【初始化仪表】建立信令连接。")
            return
        try:
            cfg = self._read_cfg()
        except ValueError as exc:
            messagebox.showwarning("输入错误", str(exc))
            return
        self.save_settings()
        _start_dt = datetime.datetime.now()   # 开始时间（用于报告末尾显示总耗时）
        self.busy = True
        self._set_busy(True)
        # 每次开始测量重置暂停状态并启用暂停按钮（测量结束时由 _poll_queue 禁用）
        self._set_status("测量中 (0/%d)..." % len(cfg["combos"]))
        # 按本次勾选的指标重建结果表列（未勾选的指标列不显示）
        self._rebuild_result_columns(cfg)
        # 清空结果表：每测完一个组合就实时插入一行
        self._clear_result_rows()
        if self.run_log:
            self.run_log.section("测试启动（组合数=%d）" % len(cfg["combos"]))
            self.run_log.kv([("IP", cfg["cmw_ip"]), ("RF口", cfg["rf"]),
                             ("频段", ",".join(band_label(b) for b in cfg["bands"])),
                             ("频段/各自带宽", band_bws_summary(
                                 cfg["bands"], cfg.get("band_bws") or {})),
                             ("带宽并集", ",".join(bw_label(w) for w in cfg["bws"])),
                             ("扫描模式", "全信道（逐 EARFCN 探底）" if cfg.get("full_scan")
                              else "低/中/高三信道"),
                             ("电源启用", cfg["psu_enable"]),
                             ("指标", ",".join(k for k, v in cfg["metrics"].items() if v)),
                             ("信道表", channel_table_source())])
            skipped = cfg.get("skipped_combos") or []
            if skipped:
                self.run_log.write("[跳过] 该频段不支持此带宽，已跳过: %s" % ", ".join(skipped))
        threading.Thread(target=self._measure_worker, args=(cfg,), daemon=True).start()

    def _measure_worker(self, cfg):
        try:
            _start_dt = datetime.datetime.now()   # 整体计时（耗时统计）
            enabled = {k for k, v in cfg["metrics"].items() if v}
            # 电流/电压/功率默认随电源记录，不再需要单独勾选 current
            want_current = bool(cfg.get("psu_enable"))
            items = []
            # 同频段同带宽上一信道的灵敏度(RS EPRE)，供下一信道热点续扫
            sens_hotspots = {}
            total = len(cfg["combos"])
            # 初始化仪表时已按某频段/带宽建立小区并注册 UE。以 init 实际建立的
            # (band,bw) 为起点（cell_actual）：同一频段+带宽内换信道走快速重同步
            # 路径；频段或带宽变化则走 try_fast_retune（自带 RMC 重配）。
            # 不能假定 init 带宽=测量首组合带宽——用户可能在 init 后改过勾选，
            # 此时首组合若直接走快速路径会带着旧 RMC（如 N50@1.4MHz）去测 TX，
            # UE 无法被调度上行 → MEV 全 INV/NCAP（真机 180539.log Run1 事故）。
            _cell_actual = getattr(self, "cell_actual", None)
            if _cell_actual and _cell_actual[0] and _cell_actual[1]:
                last_cell = _cell_actual
            else:
                first = cfg["combos"][0] if cfg["combos"] else ("", "", "", -1)
                last_cell = (first[0], first[1])
            # Pmin 档功率未配置时只提醒一次（每个组合都提醒会刷屏）
            _pmin_hint_done = False
            for idx, combo in enumerate(cfg["combos"], 1):
                band, bw, pos, earfcn = (list(combo) + ["", "", "默认", -1])[:4]
                label = "%s / %s / %s(CH%s)" % (
                    band_label(band) if band else "默认频段",
                    bw_label(bw) if bw else "默认带宽", pos, earfcn)
                self.log("[组合 %d/%d] %s" % (idx, total, label))
                self._set_status("测量中 (%d/%d): %s" % (idx, total, label))

                scpi = self.scpi
                # === 切换保护：切信道/切 STYP 前先把下行电平恢复到强信号 ===
                # 真机 log（cmw500_20261123_094316）证实：上一组合灵敏度扫描
                # 结束后电平停在灵敏度附近（-125 dBm RS EPRE，弱信号），此时
                # 切信道→切 STYP UDCH 会让 UE 掉线，且弱信号下 CALL:PSW:ACT CONN
                # 无法重新附着（PSW:STAT 一直 ON），20s 超时中止整个测量。
                # 故每次切换前先把 RS EPRE 拉回灵敏度起点（强信号），确保 UE
                # 在强信号稳定注册态下完成切换；切换后再从该电平开始扫描。
                # 仅在勾选灵敏度时生效（未勾选时电平始终为初始 rsrp 强信号）。
                if "sensitivity" in enabled:
                    try:
                        strong = float(cfg.get("sens_start", -85))
                    except (TypeError, ValueError):
                        strong = -85.0
                    cmd = "CONF:LTE:SIGN:DL:RSEPre:LEVel %.1f" % strong
                    scpi.write(cmd)
                    self.log("[SCPI] " + cmd)
                    self.log("[切换保护] 切信道前把 RS EPRE 恢复到强信号 %.1f dBm"
                             "（避免灵敏度附近弱信号切换掉线）" % strong)
                # 换频段/带宽：连接态热改参数（参考工具 18340.log 实证模式：
                # 不重启小区，UE 保持注册，~3-5s）；同频段热切绝不重发 CELL ON
                # （对已 ON 小区重发 ON = 小区整机重启 → 20s 查询黑洞 + UE 掉线，
                # 真机 191624.log 三次带宽切换 50~60s 的根源）。UE 掉线时先试
                # AT CFUN 强制重注册，再回退完整小区重建。
                # 仅换信道（同频段同带宽）走下面的快速重同步路径。
                retuning = bool(band) and (band != last_cell[0] or bw != last_cell[1])
                # FDD↔TDD 跨双工必须重建小区：duplex 只能在 CELL OFF 后切换，
                # 且 BAND 合法范围随 duplex 变化（手册 p.306），热切必被仪表拒绝。
                duplex_change = bool(band) and bool(last_cell[0]) and \
                    band_duplex(band) != band_duplex(last_cell[0])
                if duplex_change and self.log:
                    self.log("[频段快切] FDD↔TDD 跨双工切换 → 强制重建小区"
                             "（duplex 只能在 CELL OFF 后切换）")
                if retuning and duplex_change:
                    retune_cell(scpi, cfg, band, bw, earfcn, self.log)
                    # 重建后必须同步 last_cell：否则下一同频段组合仍以旧频段为
                    # 基准判 duplex_change，被重复误判为跨双工、反复重建小区
                    # （134354.log 中 B34 中/高两信道各多花约 10s 的根因）。
                    last_cell = (band, bw)
                elif retuning:
                    if not try_fast_retune(scpi, cfg, band, bw, earfcn, self.log,
                                           last_cell=last_cell):
                        # 快切后 UE 未回网：先试 AT CFUN 强制重注册（免重建），
                        # 未生效再回退完整小区重建（AT 被禁用/无口时零开销跳过）
                        _st_at = at_recover_ue(scpi, cfg, self.log)
                        if _st_at:
                            self.log("[切换] AT 强制重注册已恢复（%s），跳过小区重建" % _st_at)
                        else:
                            retune_cell(scpi, cfg, band, bw, earfcn, self.log)
                    last_cell = (band, bw)
                else:
                    if band:
                        cmd = "CONF:LTE:SIGN:BAND %s" % band
                        scpi.write(cmd); self.log("[SCPI] " + cmd)
                    if bw:
                        cmd = "CONF:LTE:SIGN:CELL:BAND:DL %s" % bw
                        scpi.write(cmd); self.log("[SCPI] " + cmd)
                    if earfcn is not None and int(earfcn) >= 0:
                        cmd = build_channel_cmd(cfg, earfcn)
                        scpi.write(cmd); self.log("[SCPI] " + cmd)
                    # 动态 EATT：同频段换信道也要按新信道中心频率补偿外部线缆损耗
                    _mhz = earfcn_to_dl_center_mhz(band, earfcn)
                    _loss = lookup_loss_mhz(cfg, _mhz) if _mhz else None
                    if _loss is not None:
                        scpi.write("CONF:LTE:SIGN:RFS:EATT:OUTP %.1f" % _loss)
                        scpi.write("CONF:LTE:SIGN:RFS:EATT:INP %.1f" % _loss)
                        self.log("[EATT] DL=%.1f MHz → 外部线缆损耗 %.1f dB（OUTP/INP 均补偿）" % (_mhz, _loss))
                    # 手册 23.2.3 (p.1681): 切换信道后连接进入
                    # "Signaling in Progress" 重同步，UE 需重新同步才可能被调度
                    # 上行发射。给足重同步时间，避免切完立即测得全 INV。
                    try:
                        chan_settle = max(0.0, min(10.0, float(cfg.get("chan_settle", 1.5))))
                    except (TypeError, ValueError):
                        chan_settle = 1.5
                    if chan_settle:
                        self.log("[重同步] 切换参数后等待 %.1fs 让 UE 在新信道"
                                 "重新同步（切信道会触发 Signaling in Progress）。" % chan_settle)
                        time.sleep(chan_settle)
                    # 配置校验：SYST:ERR? 清错误队列（对齐参考工具），及早发现
                    # -203 option missing 之类的静默配置失败
                    check_scpi_errors(scpi, self.log,
                                      "组合 %d/%d 配置后" % (idx, total))

                need_rf = bool(enabled & {"maxpower", "aclr", "evm", "freqerr"})
                need_ue = need_rf or ("sensitivity" in enabled)
                # KEEP 模式补强：同频段同带宽切信道走 else 分支，不经过
                # _apply_cell_params 的 RMC 重配；参考工具 1756.log 每次切
                # 信道后都重发 RMC:UL + RBP:UL（140/500/656 行），否则 UE
                # 虽注册 ATT 却不按 RMC 调度上行 → MEV 全 INV（140933.log）。
                # retuning 路径已由 _apply_cell_params 重发，此处不重复。
                if need_rf and band and not retuning:
                    _rb = 6 if bw == "B014" else int(BW_TO_MHZ.get(bw, 10) * 5)
                    for _cmd in ("CONF:LTE:SIGN:CONN:RMC:UL N%d,QPSK,KEEP" % _rb,
                                 "CONF:LTE:SIGN:CONN:RMC:RBP:UL LOW"):
                        scpi.write(_cmd)
                        self.log("[SCPI] " + _cmd)
                # === 测前复查注册态 ===
                # 每个组合都会重配 BAND/带宽/CHAN:DL，可能使 UE 掉线；
                # 必须确认仍处于注册/连接态才测，杜绝"没注册就测试"。
                if need_ue and not retuning:
                    try:
                        ensure_registered(scpi, self.log)
                    except ScpiError as exc:
                        self.log("[中止] 组合 %d/%d 测前 UE 未注册：%s" % (idx, total, exc))
                        self.msg_queue.put(("measure_failed", str(exc)))
                        return

                # === 测试场景（TX/RX 公用）===
                # 选 Pmax / Pmin / Pmax+Pmin 后，TX 与 RX 指标都在该功率档下测。
                mode = str(cfg.get("power_scenario", cfg.get("sens_power_mode", "high")) or "high")
                if mode not in SENS_POWER_MODES:
                    mode = "high"
                # 单场景时首个 TX 测量直接用该场景功率档；both 先测 Pmax，
                # Pmin 档 TX 随后补测（勾选灵敏度时在 low 分支内顺带测，
                # 否则在下方的独立 [功率重测] 块里单独测），两档数据齐全。
                first_pm = "low" if mode == "low" else "high"
                if need_rf:
                    self.log("[场景] 组合 %d/%d：TX 功率档=%s（%s）%s" %
                             (idx, total, SENS_POWER_LABELS.get(first_pm, first_pm),
                              first_pm,
                              "，both 场景：先测 Pmax，Pmin 档随后补测" if mode == "both" else ""))
                    set_sens_power(scpi, cfg, first_pm, self.log)

                data = {}
                if need_rf:
                    data.update(measure_tx(scpi, cfg, enabled, self.log, band=band,
                                           power_mode=first_pm))
                    if first_pm == "low":
                        data["mod_avg_low"] = data.get("mod_avg") or []
                        data["aclr_low"] = data.get("aclr") or []
                    # === 快失败：首个射频组合结果全 INV → UE 未发射，立即中止 ===
                    # 已注册却仍全 INV，说明 UE 没在上行发射；继续跑剩余组合
                    # 只会得到一堆无效值并导出误导性空报告，故立即停止、不导出。
                    if not data.get("valid", True):
                        diag = data.get("diagnosis") or \
                            diagnose_reliability(data.get("reliability"), cfg)
                        self.log("[中止] 组合 %d/%d 的 CMW 上行测量结果经 %d 次仍全部"
                                 "无效（INV/NAV）。Reliability=%s：%s\n%s" %
                                 (idx, total, data.get("attempts", 1),
                                  data.get("reliability"), data.get("reliability_text", ""), diag))
                        self.msg_queue.put(("measure_failed",
                            "组合 %d/%d：UE 已处于注册态(ATT)，但 CMW 上行测量经 %d 次"
                            "仍全部返回无效值。\n\n"
                            "Reliability 指示 = %s：%s\n\n%s\n\n"
                            "继续测量剩余组合也只会得到无效数据，已中止，未生成误导性空报告。" %
                            (idx, total, data.get("attempts", 1),
                             data.get("reliability"), data.get("reliability_text", ""), diag)))
                        return
                if want_current and self.psu:
                    try:
                        idle = read_psu_snapshot(self.psu, cfg, self.log)
                        snap = read_psu_snapshot(self.psu, cfg, self.log)
                        if snap:
                            if idle:
                                snap["idle_current"] = idle.get("current")
                            # 循环前快照：UE 处于满功率发射（TX 指标刚测完），
                            # 记作 Pmax 档电流。low/both 模式还需要 Pmin 档电流，
                            # 在 low 档灵敏度测完后再采一次（见循环内）。
                            mode = cfg.get("sens_power_mode", "high")
                            if mode == "both":
                                data["psu_high"] = snap
                                data["psu_low"] = None  # 待 low 档后采
                            elif mode == "low":
                                # 循环前先存一份作为兜底（low 档是循环内唯一档），
                                # low 档后采到 psu_low 会覆盖掉
                                data["psu"] = snap
                                data["psu_low"] = snap  # 兜底：low 档后重采前先用这份
                            else:
                                data["psu_high"] = snap
                                data["psu"] = snap
                        else:
                            self.log("[电源连接] ⚠️ 运行中断开，功耗数据跳过（射频指标不受影响）")
                            self.psu.close()
                            self.psu = None
                            self.psu_label.config(text="⚠️ 连接失败（功耗跳过）", foreground="#e67e22")
                    except Exception as e:
                        self.log("[电源连接] ⚠️ 运行中异常：%s — 功耗跳过" % type(e).__name__)
                        try:
                            self.psu.close()
                        except Exception:
                            pass
                        self.psu = None
                        self.psu_label.config(text="⚠️ 连接失败（功耗跳过）", foreground="#e67e22")
                if "sensitivity" in enabled:
                    mode = cfg.get("sens_power_mode", "high")
                    if mode not in SENS_POWER_MODES:
                        mode = "high"
                    # 按选定的功率档测量：high 只测大功率；low 只测小功率；
                    # both 先大后小，每档独立热点续扫，结果分别落两列。
                    run_modes = ["high", "low"] if mode == "both" else [mode]
                    for pm in run_modes:
                        hkey = (band, bw, pm)
                        hot = sens_hotspots.get(hkey)
                        # 首信道（无实测热点）用该 (band,bw) 的 3GPP 理论 REFSENS 作种子
                        seed = None if hot is not None else refsens_epre_for(band, bw)
                        self.log("[灵敏度] 组合 %d/%d 功率档=%s（%s）" %
                                 (idx, total, pm, SENS_POWER_LABELS.get(pm, pm)))
                        try:
                            sens = measure_sensitivity(scpi, cfg, enabled, self.log,
                                                       band=band, bw=bw, hotspot=hot,
                                                       theory_seed=seed, power_mode=pm)
                        except ScpiError as exc:
                            self.log("[中止] 组合 %d/%d 功率档=%s 灵敏度后 UE 恢复失败：%s" %
                                     (idx, total, pm, exc))
                            self.msg_queue.put(("measure_failed", str(exc)))
                            return
                        if sens:
                            # pm=="high" 永远写入 data["sensitivity"]（大功率列），
                            # pm=="low"  永远写入 data["sensitivity_low"]（小功率列）。
                            # 下游 make_wide_row 按 sens_power_mode 决定取哪一列显示：
                            #   high only → 只取 sensitivity
                            #   low  only → 只取 sensitivity_low
                            #   both      → 两列都取
                            if pm == "high":
                                data["sensitivity"] = sens
                            else:
                                data["sensitivity_low"] = sens
                            if sens.get("result_epre") is not None:
                                sens_hotspots[hkey] = sens["result_epre"]
                                self.log("[灵敏度扫描] 热点缓存 %s/%s[%s] = %.1f dBm RS EPRE，"
                                         "供同频段下一信道续扫" %
                                         (band_label(band), bw_label(bw), pm,
                                          sens["result_epre"]))
                        # low/both 模式下，low 档灵敏度测完后再采一次 PSU，
                        # 记作 Pmin 档电流（循环前快照是 Pmax 档电流）
                        if pm == "low" and want_current and self.psu:
                            try:
                                _psu_low = read_psu_snapshot(self.psu, cfg, self.log)
                                if _psu_low:
                                    if data.get("psu_high") and data["psu_high"].get("idle_current"):
                                        _psu_low["idle_current"] = data["psu_high"]["idle_current"]
                                    data["psu_low"] = _psu_low
                            except Exception:
                                pass
                        # low/both 模式下，low 档灵敏度测完 UE 仍处于 CLOop 功率态，
                        # 重测一次 TX 指标（maxpower/ACLR/EVM/freqerr），存为 low 档结果。
                        # 这样 make_wide_row 可按 RX 功率档取对应档的 TX 指标值。
                        # both 场景才需要补测 Pmin 档 TX；单 Pmin 场景首个 TX
                        # 已在该档测过，避免重复测量（见上方 first_pm）。
                        if need_rf and pm == "low" and mode == "both":
                            try:
                                # low 档弱信号扫描可能把 UE 探掉线：重测 TX 前必须确认
                                # 已注册，掉线则先重附着，避免 3 次 Underdriven 空耗重试。
                                try:
                                    _st_low = scpi.query(CONN_RECHECK_QUERY, timeout=2.0, expect=RE_PSW)
                                except Exception:
                                    _st_low = ""
                                if not conn_ready(_st_low):
                                    self.log("[功率重测] UE 未注册（%s），先重附着再测 TX" % _st_low)
                                    # 方案 A：同样启用硬掉线快速升级（重建小区后 RMC
                                    # 满 RB + 强信号正好满足随后的 TX 测量）
                                    reconnect_after_styp(
                                        scpi, self.log, cfg=cfg,
                                        retune_ctx={"band": band, "bw": bw, "earfcn": earfcn})
                                self.log("[功率重测] 组合 %d/%d：low 档灵敏度测完，UE 处于 CLOop 功率态，重测 TX 指标" % (idx, total))
                                _low_meas = measure_tx(scpi, cfg, enabled, self.log,
                                                       band=band, power_mode="low")
                                if _low_meas:
                                    if _low_meas.get("mod_avg"):
                                        data["mod_avg_low"] = _low_meas["mod_avg"]
                                    if _low_meas.get("aclr"):
                                        data["aclr_low"] = _low_meas["aclr"]
                            except Exception:
                                pass  # low 档重测失败不影响 high 档已有结果
                        # 切到下一功率档前，把电平拉回强信号，避免弱信号切档掉线
                        if pm != run_modes[-1]:
                            try:
                                _strong = float(cfg.get("sens_start", -85))
                            except (TypeError, ValueError):
                                _strong = -85.0
                            _cmd = "CONF:LTE:SIGN:DL:RSEPre:LEVel %.1f" % _strong
                            scpi.write(_cmd)
                            self.log("[SCPI] " + _cmd)
                            self.log("[切换保护] 切功率档前把 RS EPRE 恢复到强信号 %.1f dBm"
                                     "（避免灵敏度附近弱信号切换掉线）" % _strong)
                # === both 场景：Pmin 档 TX 补测（与灵敏度解耦）===
                # 原先 Pmin 档 TX 只在「勾选灵敏度」时顺带补测（在 low 档灵敏度
                # 之后，省一次功率切换）。只勾 TX 指标时就没有这一步，@Pmin 列
                # 只能拿到 Pmax 的重复值 —— 171835.log 里 Power 两列都是 21.92
                # 就是这么来的。这里补一条独立路径：只要 both 场景 + 需要 TX +
                # Pmin 档数据缺失，就在强信号下切到 Pmin 档单独测一遍。
                if (need_rf and mode == "both" and not data.get("mod_avg_low")):
                    self.log("[功率重测] 组合 %d/%d：both 场景（Pmax+Pmin），"
                             "Pmin 档 TX 指标单独补测" % (idx, total))
                    try:
                        # 提示放在下发命令之前，日志顺序才符合因果（否则会晚 4s 才出现）
                        if (not _pmin_hint_done
                                and not (cfg.get("sens_low_power") or "").strip()):
                            _pmin_hint_done = True
                            self.log("[功率重测] ⚠️ Pmin 档未指定功率 → 用 MINPower"
                                     "（UE 最小发射功率，可能低到 -45dBm 量级），CMW 有可能报"
                                     " Underdriven；若要测得到，请在「① 测试场景」的 Pmin "
                                     "输入框填一个可测功率（如 -40）。本提示每次运行只出现一次。")
                        set_sens_power(scpi, cfg, "low", self.log)
                        # 弱功率档下 UE 可能已掉线：先复查注册，掉线先重附着
                        try:
                            _st_low2 = scpi.query(CONN_RECHECK_QUERY, timeout=2.0,
                                                  expect=RE_PSW)
                        except Exception:
                            _st_low2 = ""
                        if not conn_ready(_st_low2):
                            self.log("[功率重测] UE 未注册（%s），先重附着再测 Pmin 档 TX"
                                     % _st_low2)
                            reconnect_after_styp(
                                scpi, self.log, cfg=cfg,
                                retune_ctx={"band": band, "bw": bw, "earfcn": earfcn})
                        _low_meas = measure_tx(scpi, cfg, enabled, self.log,
                                               band=band, power_mode="low")
                        if _low_meas:
                            if _low_meas.get("mod_avg"):
                                data["mod_avg_low"] = _low_meas["mod_avg"]
                            if _low_meas.get("aclr"):
                                data["aclr_low"] = _low_meas["aclr"]
                            if not _low_meas.get("valid", True):
                                self.log("[功率重测] ⚠️ Pmin 档 TX 结果无效"
                                         "（Reliability=%s：%s），@Pmin 列留空" %
                                         (_low_meas.get("reliability"),
                                          _low_meas.get("reliability_text", "")))
                        # Pmin 档电流：此刻 UE 正处于小功率发射态，采一次作数
                        if want_current and self.psu and not data.get("psu_low"):
                            try:
                                _psu_low2 = read_psu_snapshot(self.psu, cfg, self.log)
                                if _psu_low2:
                                    if data.get("psu_high") and \
                                            data["psu_high"].get("idle_current"):
                                        _psu_low2["idle_current"] = \
                                            data["psu_high"]["idle_current"]
                                    data["psu_low"] = _psu_low2
                            except Exception:
                                pass
                    except Exception as exc:
                        self.log("[功率重测] ⚠️ Pmin 档 TX 补测失败：%s，@Pmin 列留空"
                                 % exc)
                data["raw"] = data.get("raw", {})

                data["combo"] = label
                lim = {
                    "evm_lim": cfg["evm_lim"], "aclr_lim": cfg["aclr_lim"],
                    "freq_lim": cfg["freq_lim"], "sens_lim": cfg["sens_lim"],
                }
                rows = make_rows(data, lim)
                wide = make_wide_row(label, data, lim, metrics=cfg["metrics"],
                                     sens_power_mode=cfg.get("sens_power_mode", "high"),
                                     psu_enable=bool(cfg.get("psu_enable", True)))
                items.append({"combo": label, "data": data, "rows": rows, "wide": wide})

                # --- 每列独立 FAIL 判定（单元格级标红，不再整行标红）---
                try:
                    _fail_flags = _build_fail_flags(wide[:self.wide_col_count], label, cfg)
                except Exception:
                    _fail_flags = [False] * self.wide_col_count
                # 实时显示：把本组合结果行推给主线程插入结果表
                # payload = (wide_list, fail_flags_list) —— 每个单元格独立红/不红
                self.msg_queue.put(("row", (wide[:self.wide_col_count], _fail_flags)))

            # 汇总有效性：若勾选了 CMW 射频指标却全部无效 → UE 未注册/未发射
            rf_enabled = bool(enabled & {"maxpower", "aclr", "evm", "freqerr"})
            any_valid = any((it["data"].get("valid", True)) for it in items)
            all_invalid = bool(rf_enabled and items and not any_valid)
            if all_invalid:
                self.log("[失败] 全部 %d 个组合的 CMW 测量结果无效（INV/NAV）。"
                         "UE 未注册或未在上行发射，本次数据不可用。" % len(items))
            _end_dt = datetime.datetime.now()
            _secs = (_end_dt - _start_dt).total_seconds()
            result = {
                "time": _end_dt.isoformat(timespec="seconds"),
                "start_time": _start_dt.isoformat(timespec="seconds"),
                "end_time": _end_dt.isoformat(timespec="seconds"),
                "elapsed_sec": round(_secs, 1),
                "elapsed_str": _fmt_elapsed(_secs),
                "cfg": cfg, "items": items, "all_invalid": all_invalid,
            }
            os.makedirs(RESULT_DIR, exist_ok=True)
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            xlsx = os.path.join(RESULT_DIR, "LTE_%s.xlsx" % stamp)
            txt = os.path.join(RESULT_DIR, "LTE_%s.txt" % stamp)
            ok, err = export_results(result, xlsx, txt)
            result["export"] = {"xlsx": xlsx, "txt": txt, "ok": ok, "err": err}
            self.msg_queue.put(("result", result))
        except Exception:
            self.msg_queue.put(("error", traceback.format_exc()))

    def _open_config(self):
        """打开 config.json（系统默认 JSON 编辑器），方便修改 freq_loss_segments 等参数。"""
        try:
            import os as _os
            _os.startfile(CONFIG_PATH)
        except Exception as e:
            self._set_status("打开 config.json 失败: %s" % e)

    def _open_specs(self):
        """打开 specs 目录（频段表/参考灵敏度/UL 配置都在这里）。"""
        try:
            import os as _os
            _os.startfile(os.path.join(APP_DIR, "specs"))
        except Exception as e:
            self._set_status("打开 specs 目录失败: %s" % e)

    # ---------------- 自动消失提示框 ----------------
    def _show_auto_notice(self, title, text, ms=None):
        """非模态提示框：ms 毫秒后自动关闭，无需人工点击确认。

        初始化完成等提示不能阻塞主线程/中断测试；Toplevel 不 grab 焦点，
        定时器到点或用户提前点“确定”都会关闭，重复 destroy 被静默忽略。
        """
        if ms is None:
            try:
                ms = int(float(self.cfg_loaded.get("init_popup_ms", 3000)))
            except (TypeError, ValueError):
                ms = 3000
        ms = max(500, min(60000, ms))
        try:
            win = tk.Toplevel(self)
            win.title(title)
            win.transient(self)
            win.resizable(False, False)
            tk.Label(win, text=text, justify="left", wraplength=520,
                     padx=18, pady=14).pack()
            tk.Button(win, text="确定", width=8,
                      command=lambda: _close_tk_window(win)).pack(pady=(0, 12))
            win.after(ms, lambda: _close_tk_window(win))
        except Exception:
            # 任何 GUI 异常都只写日志，绝不阻塞测量流程
            self.log("[提示] %s：%s" % (title, text))

    # ---------------- 队列轮询 ----------------
    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == "log":
                    self._log_direct(payload)
                    continue
                if kind == "row":
                    # 测量结果实时逐行显示（过程消息，不解除 busy）
                    # payload = (wide_vals, fail_flags) —— fail_flags 是 bool 列表
                    try:
                        if isinstance(payload, tuple) and len(payload) == 2:
                            _wide_vals, _flags = payload
                            # 兼容旧版 payload：第二个元素是行级 tag 字符串而非 bool 列表
                            if isinstance(_flags, str):
                                _flags = [False] * len(_wide_vals)
                        else:
                            _wide_vals = payload
                            _flags = [False] * len(_wide_vals)
                        self._append_result_row(list(_wide_vals), list(_flags))
                    except Exception:
                        pass
                    continue
                if kind == "at_ports":
                    # payload = (端口名列表, 握手成功的 AT 口)；更新下拉框与选中值
                    names, found = payload
                    try:
                        cur = self.at_port_var.get().strip()
                        vals = list(dict.fromkeys(["auto"] + list(names)))
                        self.at_port_cb.config(values=vals)
                        if found and (not cur or cur.lower() == "auto"):
                            self.at_port_var.set(found)
                        self._at_port_detected = found
                    except Exception:
                        pass
                    continue
                if kind == "at_done":
                    self._at_busy = False
                    self.btn_at_send.config(state="normal")
                    continue
                if kind == "at_save":
                    # AT 发送完成后保存输入框内容（含 at_cmd/at_port），下次启动恢复
                    try:
                        self.save_settings()
                    except Exception:
                        pass
                    continue
                self.busy = False
                self._set_busy(False)
                if kind == "init_done":
                    self.initialized = True
                    self._set_status(payload)
                    self.log("[初始化] " + payload)
                    # 自动消失，无需人工点击确认（时间由 init_popup_ms 配置）
                    self._show_auto_notice("初始化完成", payload)
                elif kind == "result":
                    self._show_result(payload)
                elif kind == "measure_failed":
                    # CMW 上行测量无效导致的中止：明确报错，绝不提示"测试完成"
                    self._set_status("测量中止：CMW 上行测量无有效结果")
                    self.log("[测量中止] " + payload)
                    messagebox.showerror("测量中止 — 上行信号无效（见 Reliability 诊断）", payload)
                else:
                    self._set_status("出错，详见日志")
                    self.log("[错误]\n" + payload)
                    messagebox.showerror("操作失败", payload[-800:])
        except queue.Empty:
            pass
        except Exception as _e:
            # 兜底：任何意外异常都不能终止 poll 循环
            import traceback as _tb
            try:
                self.log("[GUI] poll_queue 意外异常（已兜底，继续轮询）: %s\n%s" %
                         (type(_e).__name__, _tb.format_exc(limit=3)))
            except Exception:
                pass
        self.after(80, self._poll_queue)

    # ---------------- Frame+Label 网格结果表（单元格级标红） ----------------
    # 每个单元格是独立 ttk.Label，FAIL 的单元格单独设红色 foreground。
    # 替代 Treeview：Treeview 只支持行级 tag 无法实现单元格级标红。

    # 冻结窗格：左上角固定区（表头行 + 左 3 列），主区双向滚动。
    # 用 4 个区域拼一张表：corner(固定) / 表头 Canvas(横向同步) /
    # 左列 Canvas(纵向同步) / 主 Canvas(双向，挂滚动条)。
    RESULT_PREFIX_COLS = 3

    def _build_result_grid(self, cfg):
        """构建四象限结果表：首行与左三列在双向滚动时保持不变。"""
        headers = wide_headers(cfg.get("metrics"),
                               cfg.get("sens_power_mode", "high"),
                               bool(cfg.get("psu_enable", True)))
        self.wide_col_count = len(headers)
        prefix_n = self.RESULT_PREFIX_COLS
        self._prefix_w = 96
        # 指标列宽：要容得下表头两行里较长的一行（如 '@Pmax (dBm)' 约 88px），
        # 否则 Tk 会把该行再折成第三行——列少时直接给 96px。
        self._metric_w = 96 if len(headers) <= 5 else 98
        _body_h = RESULT_VISIBLE_ROWS * RESULT_ROW_H
        _left_w = prefix_n * self._prefix_w

        # --- 首次：四象限 Canvas + 滚动条（之后复用） ---
        if not hasattr(self, "_res_canvas") or self._res_canvas is None:
            self._corner = ttk.Frame(self.res_frame, width=_left_w,
                                     height=RESULT_HEADER_H)
            self._hdr_canvas = tk.Canvas(self.res_frame, highlightthickness=0,
                                         height=RESULT_HEADER_H)
            self._left_canvas = tk.Canvas(self.res_frame, highlightthickness=0,
                                          width=_left_w, height=_body_h)
            self._res_canvas = tk.Canvas(self.res_frame, highlightthickness=0,
                                         height=_body_h)
            self._res_hsb = ttk.Scrollbar(self.res_frame, orient="horizontal")
            self._res_vsb = ttk.Scrollbar(self.res_frame, orient="vertical")
            # 主 Canvas 滚动 → 同步表头横滚 / 左列纵滚 / 滚动条
            self._res_canvas.configure(
                xscrollcommand=self._on_res_xscroll,
                yscrollcommand=self._on_res_yscroll)
            # 滚动条反向驱动主 Canvas（其 scrollcommand 再同步另外两个）
            self._res_hsb.configure(
                command=lambda *a: self._res_canvas.xview(*a))
            self._res_vsb.configure(
                command=lambda *a: self._res_canvas.yview(*a))
            # 布局：row0 = 冻结表头行；row1 = 数据区；row2 = 横向滚动条
            self._corner.grid(row=0, column=0, sticky="nsew")
            self._hdr_canvas.grid(row=0, column=1, sticky="nsew")
            self._left_canvas.grid(row=1, column=0, sticky="nsew")
            self._res_canvas.grid(row=1, column=1, sticky="nsew")
            self._res_vsb.grid(row=1, column=2, sticky="ns")
            self._res_hsb.grid(row=2, column=1, sticky="ew")
            # 可视区尺寸变化后重新判断是否需要横向滚动条（列少时隐藏）
            self._res_canvas.bind("<Configure>", self._on_res_canvas_config, add="+")
            self.res_frame.grid_rowconfigure(0, weight=0, minsize=RESULT_HEADER_H)
            self.res_frame.grid_rowconfigure(1, weight=0, minsize=_body_h)
            self.res_frame.grid_columnconfigure(0, weight=0, minsize=_left_w)
            self.res_frame.grid_columnconfigure(1, weight=1)
            # 滚轮：普通→纵向，Shift→横向；绑定四个区域保证任意位置都能滚
            for w in (self._corner, self._hdr_canvas,
                      self._left_canvas, self._res_canvas):
                w.bind("<MouseWheel>", self._on_res_mousewheel)
                w.bind("<Shift-MouseWheel>", self._on_res_mousewheel)

        # --- 销毁旧内部 grid 并重建 ---
        # 先删除 Canvas 里的旧 window item：只 destroy frame 会留下尺寸陈旧的
        # item，bbox("all") 仍是旧列宽，导致指标勾少后横向滚动条不消失。
        for cv in (self._res_canvas, self._left_canvas, self._hdr_canvas):
            cv.delete("all")
        for name in ("_grid_frame", "_left_frame", "_hdr_frame"):
            old = getattr(self, name, None)
            if old is not None:
                old.destroy()
        for w in self._corner.winfo_children():
            w.destroy()
        self._grid_frame = ttk.Frame(self._res_canvas)     # 指标列（数据行）
        self._left_frame = ttk.Frame(self._left_canvas)    # 左 3 列（数据行）
        self._hdr_frame = ttk.Frame(self._hdr_canvas)      # 指标列表头
        self._res_canvas.create_window((0, 0), window=self._grid_frame, anchor="nw")
        self._left_canvas.create_window((0, 0), window=self._left_frame, anchor="nw")
        self._hdr_canvas.create_window((0, 0), window=self._hdr_frame, anchor="nw")
        self._grid_row = 0

        # 表头加粗字体
        _hdr_font = tkfont.nametofont("TkDefaultFont")
        try:
            _bold = tkfont.Font(font=_hdr_font)
            _bold.configure(weight="bold")
        except Exception:
            _bold = _hdr_font
        # 冻结表头：左 3 列放 corner，其余放 hdr_frame
        # wraplength 锁定表头不撑宽列（如 "LCRB @ RBstart" 会把列撑到 101px，
        # 导致表头与数据行错开）；表头/数据四区列宽必须逐列相等。
        # 显示用 header_display()：名称一行、场景标注 + 单位一行（共 2 行）
        # （'Power@Pmax\n(dBm)' → 'Power\n@Pmax (dBm)'）。
        disp = [header_display(h) for h in headers]
        for ci, h in enumerate(disp[:prefix_n]):
            ttk.Label(self._corner, text=h, font=_bold, anchor="center",
                      justify="center", relief="solid", borderwidth=1,
                      wraplength=self._prefix_w - 6,
                      background="#e8e8e8").grid(
                row=0, column=ci, sticky="nsew", padx=1, pady=1)
        for ci, h in enumerate(disp[prefix_n:]):
            ttk.Label(self._hdr_frame, text=h, font=_bold, anchor="center",
                      justify="center", relief="solid", borderwidth=1,
                      wraplength=self._metric_w - 6,
                      background="#e8e8e8").grid(
                row=0, column=ci, sticky="nsew", padx=1, pady=1)
        # 表头高度按实际行数自适应（场景标注 2 行，窗口字体放大时也不会裁字）
        _nlines = max((len(x.split("\n")) for x in disp), default=2)
        hdr_h = max(RESULT_HEADER_H, _nlines * RESULT_HEADER_LINE_H + 6)
        try:
            self._hdr_canvas.configure(height=hdr_h)
            self.res_frame.grid_rowconfigure(0, weight=0, minsize=hdr_h)
        except Exception:
            pass
        # 列宽 / 行高固定，保证四个区域逐格对齐
        nm = max(0, len(headers) - prefix_n)
        for fr, count, w in ((self._corner, prefix_n, self._prefix_w),
                             (self._left_frame, prefix_n, self._prefix_w),
                             (self._hdr_frame, nm, self._metric_w),
                             (self._grid_frame, nm, self._metric_w)):
            for ci in range(count):
                fr.columnconfigure(ci, minsize=w, weight=0)
        for fr in (self._corner, self._hdr_frame):
            fr.rowconfigure(0, minsize=hdr_h, weight=0)
        # scrollregion 跟随内部 frame 尺寸
        for fr, cv in ((self._grid_frame, self._res_canvas),
                       (self._left_frame, self._left_canvas),
                       (self._hdr_frame, self._hdr_canvas)):
            def _cfg(_evt, _cv=cv):
                bb = _cv.bbox("all")
                if bb:
                    _cv.configure(scrollregion=bb)
                # 行增删会改变内容宽度，需重判横向滚动条
                self.after_idle(self._update_res_hsb)
            fr.bind("<Configure>", _cfg)
        # 列数变化后立刻判断横向滚动条；首次渲染后再判一次（窗口尚未映射时宽度为1）
        self._sync_res_hsb_now()
        self.after_idle(self._update_res_hsb)
        # 重建列会清空数据行，纵向位置回到顶部
        self.after_idle(self._scroll_res_to_top)

    def _on_res_xscroll(self, first, last):
        self._res_hsb.set(first, last)
        self._hdr_canvas.xview_moveto(float(first))

    def _on_res_yscroll(self, first, last):
        self._res_vsb.set(first, last)
        self._left_canvas.yview_moveto(float(first))

    def _scroll_res_to_latest(self):
        """结果行更新后自动滚到最新一行（冻结左列由 yscrollcommand 同步）。

        放在 after_idle 执行，确保新行的 <Configure> 已把 scrollregion 撑大，
        否则 yview_moveto(1.0) 会按旧的可视范围计算而滚不到底部。
        """
        cv = getattr(self, "_res_canvas", None)
        if cv is None:
            return
        try:
            cv.yview_moveto(1.0)
        except Exception:
            pass

    def _scroll_res_to_top(self):
        """清空数据行后回到顶部。"""
        cv = getattr(self, "_res_canvas", None)
        if cv is None:
            return
        try:
            cv.yview_moveto(0.0)
        except Exception:
            pass

    def _on_res_canvas_config(self, _evt=None):
        """可视区尺寸变化后决定是否显示横向滚动条（列少时能完整显示就不显示）。"""
        self.after_idle(self._update_res_hsb)

    def _update_res_hsb(self):
        """指标列总宽 > 可视宽度才显示横向滚动条；否则隐藏并归零横向偏移。

        勾选指标后、出数据前主区内部 frame 宽度只有 1px（没有子控件），
        因此列总宽必须取「主区 bbox」与「表头 bbox」的较大值，
        这样横向滚动条（左右滑）会随勾选即时出现，不必等开始测试。
        """
        cv = getattr(self, "_res_canvas", None)
        hsb = getattr(self, "_res_hsb", None)
        if cv is None or hsb is None:
            return
        try:
            bb = cv.bbox("all")
            bb_hdr = self._hdr_canvas.bbox("all")
            content_w = max(bb[2] if bb else 0, bb_hdr[2] if bb_hdr else 0)
            view_w = cv.winfo_width()
            if view_w <= 1 or content_w <= 0:
                return
            need = content_w > view_w + 2
            shown = bool(hsb.winfo_ismapped())
            if need and not shown:
                hsb.grid(row=2, column=1, sticky="ew")
            elif not need and shown:
                hsb.grid_remove()
                # 隐藏滚动条时回到最左，避免停在不可见的滚动位置
                cv.xview_moveto(0.0)
                self._hdr_canvas.xview_moveto(0.0)
            # 主区 scrollregion 至少覆盖表头宽度：空表时也能滚到最后几列，
            # 且与表头 Canvas 使用同一比例，横滚时首行与数据行不错位。
            w = max(bb[2] if bb else 0, bb_hdr[2] if bb_hdr else 0)
            h = max(bb[3] if bb else 0, 1)
            cv.configure(scrollregion=(0, 0, w, h))
        except Exception:
            pass

    def _sync_res_hsb_now(self):
        """列重建后立即按新内容宽度刷新横向滚动条可见性。"""
        try:
            self._res_canvas.update_idletasks()
            self._update_res_hsb()
        except Exception:
            pass

    def _on_res_mousewheel(self, event):
        """Shift+滚轮 → 水平滚动；普通滚轮 → 垂直滚动（主 Canvas 同步冻结区）。"""
        _n = int(-event.delta / 120)
        if event.state & 0x1:  # Shift 按下：横向看更多指标列
            self._res_canvas.xview_scroll(_n, "units")
        else:
            self._res_canvas.yview_scroll(_n, "units")

    def _append_result_row(self, wide_vals, fail_flags):
        """追加一行：左 3 列进冻结区，其余进主区；FAIL 单元格单独标红。"""
        if not hasattr(self, "_grid_frame") or self._grid_frame is None:
            return
        r = self._grid_row
        self._grid_row += 1
        prefix_n = self.RESULT_PREFIX_COLS
        n = min(len(wide_vals), self.wide_col_count)
        for ci in range(n):
            text = str(wide_vals[ci]) if wide_vals[ci] is not None else ""
            is_fail = bool(ci < len(fail_flags) and fail_flags[ci])
            fg = "#d32f2f" if is_fail else "#333333"
            if ci < prefix_n:
                parent, col = self._left_frame, ci
            else:
                parent, col = self._grid_frame, ci - prefix_n
            _wl = (self._prefix_w if ci < prefix_n else self._metric_w) - 6
            ttk.Label(parent, text=text, foreground=fg,
                      anchor="center", justify="center", wraplength=_wl,
                      relief="solid", borderwidth=1).grid(
                row=r, column=col, sticky="nsew", padx=1, pady=1)
        for fr in (self._left_frame, self._grid_frame):
            fr.rowconfigure(r, minsize=RESULT_ROW_H, weight=0)
        # 新行插入后自动滚到最新一行，无需手动拖动上下滑块
        self.after_idle(self._scroll_res_to_latest)

    def _clear_result_rows(self):
        """清空数据行（保留冻结表头与左列区结构）。"""
        if not hasattr(self, "_grid_frame") or self._grid_frame is None:
            return
        for fr in (self._left_frame, self._grid_frame):
            for w in fr.winfo_children():
                w.destroy()
        self._grid_row = 0
        self.after_idle(self._scroll_res_to_top)

    def _rebuild_result_columns(self, cfg):
        """按本次勾选指标 + 灵敏度功率档重建结果表列。"""
        self._build_result_grid(cfg)

    def _show_result(self, result):
        self.last_result = result
        # 结果行已在测量过程中实时逐行插入，这里不再清空重填（避免闪烁）
        exp = result.get("export") or {}
        n = len(result["items"])

        # 耗时信息写入状态栏（结果区不再预留空 Label 行，去掉红框所示空白）
        _el = result.get("elapsed_str") or ""
        _elapsed_text = ""
        if _el:
            _elapsed_text = "  总耗时 %s（%.1fs）  平均 %.1fs/组合" % (
                _el, result.get("elapsed_sec", 0),
                (result.get("elapsed_sec", 0) / n) if n else 0)
            try:
                self.elapsed_label.config(text=_elapsed_text.strip())
            except Exception:
                pass

        if exp.get("ok"):
            self._set_status("测量完成：%d 组合，报告已生成 %s / %s%s" %
                    (n, exp.get("xlsx"), exp.get("txt"), _elapsed_text))
            self.log("[导出] XLSX: %s" % exp.get("xlsx"))
            self.log("[导出] TXT : %s" % exp.get("txt"))
            if exp.get("err"):
                self.log("[导出提示] %s" % exp["err"])
            if result.get("all_invalid"):
                # 数据全无效时绝不能提示"测试完成"，避免误导用户
                self._set_status("测量失败：CMW 上行测量全为无效值")
                # 取首个组合的 Reliability 诊断，给出精确原因
                first = result["items"][0]["data"] if result["items"] else {}
                diag = first.get("diagnosis") or diagnose_reliability(
                    first.get("reliability"), result.get("cfg") or {})
                messagebox.showerror(
                    "测量失败 — 上行信号无效（见 Reliability 诊断）",
                    "%d 个组合的 CMW 上行测量结果全部为无效值（INV/NAV），"
                    "本次数据不可用，未据此判定指标。\n\n"
                    "Reliability=%s：%s\n\n%s\n\n"
                    "（报告已导出仅供排查：%s）" %
                    (n, first.get("reliability"),
                     first.get("reliability_text", ""), diag, exp.get("txt")))
            else:
                # 测试完成提示（第 3 项需求）
                messagebox.showinfo(
                    "测试完成",
                    "全部 %d 个组合（频段×带宽×低/中/高信道）测量完成。\n\n"
                    "Excel: %s\nTXT: %s\n\n结果已显示在上方表格（指标按列）。" %
                    (n, exp.get("xlsx"), exp.get("txt")))
        else:
            self._set_status("测量完成，但导出失败: %s" % exp.get("err", ""))
            messagebox.showwarning("测量完成但导出失败",
                                   "测量已完成，但报告导出失败：\n%s" % exp.get("err", ""))
        if self.run_log:
            self.run_log.section("测量结果：%d 组合" % n)
            for it in result["items"]:
                self.run_log.write(" | ".join(str(x) for x in it.get("wide", [])))

    # ---------------- 关闭 ----------------
    def _on_close(self):
        self.save_settings()
        self.disconnect()
        if self.run_log:
            self.run_log.write("工具关闭")
            self.run_log.close()
        self.destroy()


# 变量名 -> 变量对象映射（供 _read_cfg 数值校验）
_VAR_MAP = {
    "enp": "enp_var", "eatt": "eatt_var", "umar": "umar_var",
    "rsrp": "rsrp_var", "pmax": "pmax_var",
    "sens_start": "sens_start_var", "sens_stop": "sens_stop_var",
    "sens_step": "sens_step_var", "sens_near": "sens_near_var",
    "sens_coarse": "sens_coarse_var",
    "sens_hotspot_margin": "sens_hotspot_margin_var",
    "sens_hotspot_span": "sens_hotspot_span_var",
    "sens_throughput": "sens_throughput_var",
    "bler_timeout": "bler_timeout_var", "tpc_settle": "tpc_settle_var",
}


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        if getattr(sys, "frozen", False):
            # windowed exe 无控制台：把自检输出写到 exe 旁 selftest_result.txt
            import io as _io
            _buf = _io.StringIO()
            _old = sys.stdout
            sys.stdout = _buf
            _ok = True
            try:
                run_selftest()
            except SystemExit as _e:
                _ok = _e.code in (0, None)
            except Exception:
                import traceback as _tb
                _tb.print_exc(file=_buf)
                _ok = False
            finally:
                sys.stdout = _old
            try:
                with open(os.path.join(APP_DIR, "selftest_result.txt"),
                          "w", encoding="utf-8") as _f:
                    _f.write(_buf.getvalue() or
                             ("selftest %s" % ("OK" if _ok else "FAIL")))
            except Exception:
                pass
            sys.exit(0 if _ok else 1)
        run_selftest()
    elif "--demo" in sys.argv:
        run_demo()
    else:
        main()
