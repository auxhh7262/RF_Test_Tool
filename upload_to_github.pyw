# -*- coding: utf-8 -*-
"""RF_Test_Tool 上传助手（图形界面，无控制台窗口）。

用途：把本项目的必要文件上传到 GitHub，且不误传运行数据。
铁律：
  1) 只暂存 git 已跟踪文件的改动 + 未跟踪但"安全"的新文件；
  2) config.json（含真实仪表 IP）、exe、logs/、result/、state/、PDF 手册等
     永远不进暂存区；即使已误暂存，提交前还会再拦一次；
  3) 提交和上传都必须点击按钮，工具不会自动 push；
  4) 上传可走代理：默认使用浏览器/系统保存的代理，也可切换环境变量、自定义或直连。
     代理设置只写本地 upload_settings.json，不上传。

运行：双击本文件（已关联 pythonw），或
      "D:\\Program Files\\Python\\Python314\\pythonw.exe" upload_to_github.pyw
"""
import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk

ROOT = os.path.dirname(os.path.abspath(__file__))
SETTINGS = os.path.join(ROOT, "upload_settings.json")

# 绝不上传：本地实跑配置（有仪表 IP）、运行数据、打包产物、版权手册、缓存
BLOCKED_EXACT = {
    "config.json", "selftest_result.txt", "upload_settings.json",
}
BLOCKED_PREFIX = (
    "logs/", "result/", "state/", "dist/", "build/", ".venv", ".git/",
    ".workbuddy/",
)
BLOCKED_SUFFIX = (
    ".exe", ".pdf", ".pyc", ".pyd", ".pdb", ".zip", ".7z",
)
BLOCKED_PARTS = (".bak_",)

PROXY_MODES = (
    "浏览器保存的代理（推荐）",
    "自动（环境变量优先，其次浏览器）",
    "自定义",
    "不使用代理",
)


def is_safe(path):
    """判断一个相对路径是否允许上传。"""
    p = path.replace("\\", "/").lstrip("./")
    low = p.lower()
    if low in BLOCKED_EXACT:
        return False
    if low.startswith(BLOCKED_PREFIX):
        return False
    if low.endswith(BLOCKED_SUFFIX):
        return False
    if any(part in p for part in BLOCKED_PARTS):
        return False
    return True


def normalize_proxy(text):
    """把注册表里的代理字符串整理成 git 可用的 URL。"""
    s = (text or "").strip()
    if not s:
        return ""
    # 注册表常见格式：http=host:port;https=host:port
    if "=" in s or ";" in s:
        parts = {}
        for item in s.split(";"):
            if "=" in item:
                k, v = item.split("=", 1)
                parts[k.strip().lower()] = v.strip()
        s = parts.get("https") or parts.get("http") or ""
    if not s:
        return ""
    if "://" not in s:
        s = "http://" + s
    return s


def browser_proxy():
    """读取 Edge/IE 共用的 Windows 用户级代理，返回 (url, enabled)。"""
    try:
        import winreg
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Internet Settings",
        ) as key:
            try:
                enabled = bool(winreg.QueryValueEx(key, "ProxyEnable")[0])
            except OSError:
                enabled = False
            try:
                server = winreg.QueryValueEx(key, "ProxyServer")[0]
            except OSError:
                server = ""
        return normalize_proxy(server), enabled
    except Exception:
        return "", False


class Uploader(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.title("RF_Test_Tool 上传助手")
        self.geometry("900x640")
        self.resizable(True, True)
        self.q = queue.Queue()
        self.busy = False
        self.pending_push = None

        top = ttk.Frame(self, padding=(10, 8))
        top.pack(fill="x")
        ttk.Label(top, text="提交说明：").grid(row=0, column=0, sticky="w")
        self.msg = ttk.Entry(top, width=72)
        self.msg.grid(row=0, column=1, sticky="ew", padx=6)
        top.columnconfigure(1, weight=1)

        pf = ttk.LabelFrame(self, text="上传代理", padding=(10, 4))
        pf.pack(fill="x")
        self.proxy_mode = ttk.Combobox(pf, values=PROXY_MODES, state="readonly", width=32)
        self.proxy_mode.grid(row=0, column=0, padx=3)
        ttk.Label(pf, text="自定义地址：").grid(row=0, column=1, sticky="w")
        self.proxy_custom = ttk.Entry(pf, width=40)
        self.proxy_custom.grid(row=0, column=2, sticky="ew", padx=3)
        self.b_proxy_test = ttk.Button(pf, text="测试连接", command=self.test_proxy)
        self.b_proxy_test.grid(row=0, column=3, padx=3)
        self.proxy_info = ttk.Label(pf, text="", foreground="#666")
        self.proxy_info.grid(row=1, column=0, columnspan=4, sticky="w", padx=3)
        pf.columnconfigure(2, weight=1)
        self.load_settings()
        self.proxy_mode.bind("<<ComboboxSelected>>", lambda _e: self.save_settings())
        self.proxy_custom.bind("<FocusOut>", lambda _e: self.save_settings())

        btns = ttk.Frame(self, padding=(10, 4))
        btns.pack(fill="x")
        self.b_refresh = ttk.Button(btns, text="刷新状态", command=self.refresh)
        self.b_diff = ttk.Button(btns, text="查看差异", command=self.diff)
        self.b_stage = ttk.Button(btns, text="暂存安全文件", command=self.stage)
        self.b_commit = ttk.Button(btns, text="仅提交", command=lambda: self.commit(False))
        self.b_push = ttk.Button(btns, text="提交并上传", command=lambda: self.commit(True))
        self.b_clear = ttk.Button(btns, text="清空日志", command=self.clear_log)
        for i, b in enumerate((self.b_refresh, self.b_diff, self.b_stage,
                               self.b_commit, self.b_push, self.b_clear)):
            b.grid(row=0, column=i, padx=3)

        mid = ttk.Frame(self, padding=(10, 4))
        mid.pack(fill="both", expand=True)
        self.log = tk.Text(mid, height=22, wrap="none", font=("Consolas", 10))
        sc = ttk.Scrollbar(mid, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=sc.set)
        self.log.pack(side="left", fill="both", expand=True)
        sc.pack(side="right", fill="y")

        foot = ttk.Label(
            self,
            text="仓库目录：%s    |    不上传 config.json / exe / logs / result / state / PDF / upload_settings.json"
                 % ROOT,
            padding=(10, 6),
        )
        foot.pack(fill="x")

        self.after(120, self._poll)
        self._run(["git", "status", "--short", "--branch"])

    # ---------- 代理 ----------
    def load_settings(self):
        data = {}
        try:
            with open(SETTINGS, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception:
            pass
        mode = data.get("proxy_mode", PROXY_MODES[0])
        self.proxy_mode.set(mode if mode in PROXY_MODES else PROXY_MODES[0])
        self.proxy_custom.delete(0, "end")
        self.proxy_custom.insert(0, data.get("proxy_custom", ""))
        detected, enabled = browser_proxy()
        self.proxy_info.configure(
            text="检测到浏览器代理：%s（%s）"
                 % (detected or "未设置",
                    "已启用" if enabled else "当前未启用，仍可强制使用")
        )

    def save_settings(self):
        try:
            with open(SETTINGS, "w", encoding="utf-8") as fh:
                json.dump({"proxy_mode": self.proxy_mode.get(),
                           "proxy_custom": self.proxy_custom.get()},
                          fh, ensure_ascii=False, indent=2)
        except Exception as exc:
            self.put("[警告] 代理设置保存失败：%s\n" % exc)

    def resolve_proxy(self):
        mode = self.proxy_mode.get()
        if mode == PROXY_MODES[3]:
            return ""
        if mode == PROXY_MODES[2]:
            return normalize_proxy(self.proxy_custom.get())
        detected, enabled = browser_proxy()
        if mode == PROXY_MODES[0]:
            return detected
        for key in ("HTTPS_PROXY", "https_proxy", "HTTP_PROXY", "http_proxy"):
            value = normalize_proxy(os.environ.get(key, ""))
            if value:
                return value
        return detected if enabled else ""

    def test_proxy(self):
        if self.busy:
            return
        self.put("当前代理：%s\n" % (self.resolve_proxy() or "不使用代理"))
        self._run(["git", "ls-remote", "origin", "HEAD"])

    # ---------- 日志 / 运行状态 ----------
    def clear_log(self):
        self.log.delete("1.0", "end")

    def put(self, text):
        self.log.insert("end", text)
        self.log.see("end")

    def _set_busy(self, busy):
        self.busy = busy
        for b in (self.b_refresh, self.b_diff, self.b_stage,
                  self.b_commit, self.b_push, self.b_clear, self.b_proxy_test):
            b.configure(state="disabled" if busy else "normal")

    def _run(self, args):
        """在后台线程执行 git，结果放回队列，避免界面卡死。"""
        self._set_busy(True)
        original = list(args)
        proxy = self.resolve_proxy() if original and original[0] == "git" else ""
        if proxy:
            args = ["git", "-c", "http.proxy=" + proxy,
                    "-c", "https.proxy=" + proxy] + original[1:]
            self.put("\n[代理 %s]\n" % proxy)
        self.put("\n> " + " ".join(original) + "\n")

        def work():
            try:
                env = dict(os.environ)
                env["GIT_PAGER"] = "cat"
                env["LC_ALL"] = "C.UTF-8"
                p = subprocess.run(args, cwd=ROOT, env=env,
                                   stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                   timeout=300)
                out = (p.stdout or b"").decode("utf-8", "replace")
                err = (p.stderr or b"").decode("utf-8", "replace")
                self.q.put((p.returncode, out, err, original, proxy))
            except Exception as exc:
                self.q.put((-1, "", str(exc), original, proxy))

        threading.Thread(target=work, daemon=True).start()

    def _poll(self):
        try:
            while True:
                code, out, err, args, proxy = self.q.get_nowait()
                if out:
                    self.put(out if out.endswith("\n") else out + "\n")
                if err:
                    self.put("[stderr] " + err + "\n")
                if code != 0:
                    self.put("[失败] exit=%s\n" % code)
                    if args[:2] == ["git", "push"]:
                        if proxy:
                            self.put("提示：代理 %s 可能不可用，可点【测试连接】"
                                     "或在代理设置里切换。\n" % proxy)
                        else:
                            self.put("提示：认证失败请先在 Git Credential Manager 登录 GitHub，"
                                     "或使用 Personal Access Token。\n")
                self._set_busy(False)
                # 提交成功后才推送，避免 commit 失败却推了旧内容
                if code == 0 and args[:2] == ["git", "commit"] and self.pending_push:
                    nxt = self.pending_push
                    self.pending_push = None
                    self._run(nxt)
        except queue.Empty:
            pass
        self.after(120, self._poll)

    # ---------- 按钮动作 ----------
    def refresh(self):
        self._run(["git", "status", "--short", "--branch"])

    def diff(self):
        self._run(["git", "diff", "--stat", "HEAD"])

    def staged_names(self):
        p = subprocess.run(["git", "diff", "--cached", "--name-only"], cwd=ROOT,
                           stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return [x.strip() for x in p.stdout.decode("utf-8", "replace").splitlines()
                if x.strip()]

    def stage(self):
        # -u 只更新已跟踪文件；新文件逐条过白名单，绝不 git add -A
        self._run(["git", "add", "-u"])
        names = subprocess.run(["git", "ls-files", "--others", "--exclude-standard"],
                               cwd=ROOT, stdout=subprocess.PIPE).stdout
        new_files = [x.strip() for x in names.decode("utf-8", "replace").splitlines()
                     if x.strip()]
        safe = [p for p in new_files if is_safe(p)]
        skipped = [p for p in new_files if not is_safe(p)]
        if skipped:
            self.put("已跳过不安全文件：%s\n" % ", ".join(skipped))
        if safe:
            self._run(["git", "add", "--"] + safe)

    def commit(self, push):
        if self.busy:
            return
        msg = self.msg.get().strip()
        if not msg:
            messagebox.showwarning("缺少提交说明", "请先填写提交说明。")
            return

        staged = self.staged_names()
        bad = [p for p in staged if not is_safe(p)]
        if bad:
            messagebox.showerror(
                "拒绝提交",
                "暂存区含禁止上传的文件：\n  %s\n\n请从暂存区移除后重试。" % "\n  ".join(bad))
            return
        if not staged:
            messagebox.showinfo("无改动", "暂存区没有文件，请先点【暂存安全文件】。")
            return
        self.put("待提交文件（%d 个）：%s\n" % (len(staged), ", ".join(staged)))
        if not messagebox.askyesno(
                "确认提交",
                "将提交 %d 个文件%s：\n\n%s\n\n继续？"
                % (len(staged), "并上传到 GitHub" if push else "",
                   "\n".join(staged))):
            return

        # 先提交，成功后再 push（避免 commit 失败仍然推送旧内容）
        self.pending_push = None
        self._run(["git", "commit", "-m", msg])
        if push:
            br = subprocess.run(["git", "rev-parse", "--abbrev-ref", "HEAD"],
                                cwd=ROOT, stdout=subprocess.PIPE).stdout
            branch = br.decode("utf-8", "replace").strip() or "main"
            self.pending_push = ["git", "push", "-u", "origin", branch]


if __name__ == "__main__":
    if not os.path.isdir(os.path.join(ROOT, ".git")):
        messagebox.showerror("找不到仓库", "本目录不是 git 仓库：%s" % ROOT)
        sys.exit(1)
    Uploader().mainloop()
