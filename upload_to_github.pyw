# -*- coding: utf-8 -*-
"""RF_Test_Tool 上传助手（图形界面，无控制台窗口）。

用途：把本项目的必要文件上传到 GitHub，且不误传运行数据。
铁律：
  1) 只暂存 git 已跟踪文件的改动 + 未跟踪但"安全"的新文件；
  2) config.json（含真实仪表 IP）、exe、logs/、result/、state/、PDF 手册等
     永远不进暂存区；即使已误暂存，提交前还会再拦一次；
  3) 提交和上传都必须点击按钮，工具不会自动 push。

运行：双击本文件（已关联 pythonw），或
      "D:\\Program Files\\Python\\Python314\\pythonw.exe" upload_to_github.pyw
"""
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk

ROOT = os.path.dirname(os.path.abspath(__file__))

# 绝不上传：本地实跑配置（有仪表 IP）、运行数据、打包产物、版权手册、缓存
BLOCKED_EXACT = {
    "config.json", "selftest_result.txt",
}
BLOCKED_PREFIX = (
    "logs/", "result/", "state/", "dist/", "build/", ".venv", ".git/",
    ".workbuddy/",
)
BLOCKED_SUFFIX = (
    ".exe", ".pdf", ".pyc", ".pyd", ".pdb", ".zip", ".7z",
)
BLOCKED_PARTS = (".bak_",)


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


class Uploader(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.title("RF_Test_Tool 上传助手")
        self.geometry("820x600")
        self.resizable(True, True)
        self.q = queue.Queue()
        self.busy = False

        top = ttk.Frame(self, padding=(10, 8))
        top.pack(fill="x")
        ttk.Label(top, text="提交说明：").grid(row=0, column=0, sticky="w")
        self.msg = ttk.Entry(top, width=72)
        self.msg.grid(row=0, column=1, sticky="ew", padx=6)
        top.columnconfigure(1, weight=1)

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
            text="仓库目录：%s    |    不会上传 config.json / exe / logs / result / state / PDF"
                 % ROOT,
            padding=(10, 6),
        )
        foot.pack(fill="x")

        self.after(120, self._poll)
        self._run(["git", "status", "--short", "--branch"])

    # ---------- 日志 / 运行状态 ----------
    def clear_log(self):
        self.log.delete("1.0", "end")

    def put(self, text):
        self.log.insert("end", text)
        self.log.see("end")

    def _set_busy(self, busy):
        self.busy = busy
        for b in (self.b_refresh, self.b_diff, self.b_stage,
                  self.b_commit, self.b_push):
            b.configure(state="disabled" if busy else "normal")

    def _run(self, args):
        """在后台线程执行 git，结果放回队列，避免界面卡死。"""
        self._set_busy(True)
        self.put("\n> " + " ".join(args) + "\n")

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
                self.q.put((p.returncode, out, err, args))
            except Exception as exc:
                self.q.put((-1, "", str(exc), args))

        threading.Thread(target=work, daemon=True).start()

    def _poll(self):
        try:
            while True:
                code, out, err, args = self.q.get_nowait()
                if out:
                    self.put(out if out.endswith("\n") else out + "\n")
                if err:
                    self.put("[stderr] " + err + "\n")
                if code != 0:
                    self.put("[失败] exit=%s\n" % code)
                    if args[:2] == ["git", "push"]:
                        self.put("提示：认证失败请先在 Git Credential Manager 登录 GitHub，"
                                 "或使用 Personal Access Token。\n")
                self._set_busy(False)
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
        self._run(["git", "commit", "-m", msg])
        if push:
            br = subprocess.run(["git", "rev-parse", "--abbrev-ref", "HEAD"],
                                cwd=ROOT, stdout=subprocess.PIPE).stdout
            branch = br.decode("utf-8", "replace").strip() or "main"
            self._run(["git", "push", "-u", "origin", branch])


if __name__ == "__main__":
    if not os.path.isdir(os.path.join(ROOT, ".git")):
        messagebox.showerror("找不到仓库", "本目录不是 git 仓库：%s" % ROOT)
        sys.exit(1)
    Uploader().mainloop()
