# -*- coding: utf-8 -*-
"""
定时任务执行器（SKU 1）
对应需求：猪八戒「Python定时请求小工具」（苏州 11-05 单同款：自定义间隔循环请求+参数动态生成+日志）
功能：
  1. 配置化定时 HTTP 请求（间隔秒/分钟可选）
  2. URL/参数/请求头模板化
  3. 失败自动重试（次数可配）
  4. 日志落盘（logs/ 目录按天滚动）
  5. GUI：启动/停止/实时日志/状态栏
打包：pyinstaller --onefile --noconsole
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import requests
import threading
import time
import json
import os
import sys
from datetime import datetime

APP_TITLE = "定时任务执行器 v1.0"
CONFIG_NAME = "config.json"
LOG_DIR = "logs"


class App:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("720x560")
        root.minsize(640, 480)

        self.running = False
        self.thread = None
        self.success = 0
        self.fail = 0
        self.retry = 0

        self._build_ui()
        self._load_config()

    # ---------- UI ----------
    def _build_ui(self):
        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)

        # URL 行
        row1 = ttk.Frame(frm)
        row1.pack(fill=tk.X, pady=3)
        ttk.Label(row1, text="请求地址：", width=12).pack(side=tk.LEFT)
        self.url_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.url_var).pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 方法+间隔
        row2 = ttk.Frame(frm)
        row2.pack(fill=tk.X, pady=3)
        ttk.Label(row2, text="请求方法：", width=12).pack(side=tk.LEFT)
        self.method_var = tk.StringVar(value="GET")
        ttk.Combobox(row2, textvariable=self.method_var, values=["GET", "POST"], width=8, state="readonly").pack(side=tk.LEFT)
        ttk.Label(row2, text="    执行间隔：", width=12).pack(side=tk.LEFT)
        self.interval_var = tk.StringVar(value="60")
        ttk.Spinbox(row2, from_=1, to=86400, textvariable=self.interval_var, width=8).pack(side=tk.LEFT)
        ttk.Label(row2, text="秒").pack(side=tk.LEFT)

        # 失败重试
        row3 = ttk.Frame(frm)
        row3.pack(fill=tk.X, pady=3)
        ttk.Label(row3, text="失败重试：", width=12).pack(side=tk.LEFT)
        self.retry_var = tk.StringVar(value="3")
        ttk.Spinbox(row3, from_=0, to=20, textvariable=self.retry_var, width=8).pack(side=tk.LEFT)
        ttk.Label(row3, text="次    请求超时：").pack(side=tk.LEFT)
        self.timeout_var = tk.StringVar(value="10")
        ttk.Spinbox(row3, from_=1, to=120, textvariable=self.timeout_var, width=6).pack(side=tk.LEFT)
        ttk.Label(row3, text="秒").pack(side=tk.LEFT)

        # 参数模板（POST body）
        row4 = ttk.Frame(frm)
        row4.pack(fill=tk.X, pady=3)
        ttk.Label(row4, text="参数模板：", width=12).pack(side=tk.LEFT, anchor=tk.N)
        self.body_txt = tk.Text(row4, height=4)
        self.body_txt.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tip = ttk.Label(frm, text="提示：参数支持 {timestamp} {date} {rand} 动态变量；GET 请求参数留空即可", foreground="#888")
        tip.pack(anchor=tk.W)

        # 按钮
        row5 = ttk.Frame(frm)
        row5.pack(fill=tk.X, pady=6)
        self.start_btn = ttk.Button(row5, text="▶ 启动", command=self.start)
        self.start_btn.pack(side=tk.LEFT, padx=4)
        self.stop_btn = ttk.Button(row5, text="■ 停止", command=self.stop, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=4)
        ttk.Button(row5, text="保存配置", command=self.save_config).pack(side=tk.LEFT, padx=4)
        ttk.Button(row5, text="打开日志目录", command=self.open_logs).pack(side=tk.LEFT, padx=4)
        ttk.Button(row5, text="测试一次", command=self.test_once).pack(side=tk.LEFT, padx=4)

        # 状态栏
        row6 = ttk.Frame(frm)
        row6.pack(fill=tk.X, pady=3)
        self.status_var = tk.StringVar(value="状态：未运行")
        ttk.Label(row6, textvariable=self.status_var, font=("微软雅黑", 10, "bold")).pack(side=tk.LEFT)
        self.count_var = tk.StringVar(value="成功 0 | 失败 0")
        ttk.Label(row6, textvariable=self.count_var).pack(side=tk.RIGHT)

        # 日志框
        logf = ttk.LabelFrame(frm, text="运行日志", padding=4)
        logf.pack(fill=tk.BOTH, expand=True, pady=4)
        self.log_txt = tk.Text(logf, height=12, state=tk.DISABLED, font=("Consolas", 9))
        self.log_txt.pack(fill=tk.BOTH, expand=True)

    # ---------- 配置 ----------
    def _config_path(self):
        return os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), CONFIG_NAME)

    def _load_config(self):
        try:
            with open(self._config_path(), encoding="utf-8") as f:
                cfg = json.load(f)
            self.url_var.set(cfg.get("url", ""))
            self.method_var.set(cfg.get("method", "GET"))
            self.interval_var.set(str(cfg.get("interval", 60)))
            self.retry_var.set(str(cfg.get("retry", 3)))
            self.timeout_var.set(str(cfg.get("timeout", 10)))
            self.body_txt.delete("1.0", tk.END)
            self.body_txt.insert("1.0", cfg.get("body", ""))
        except Exception:
            pass

    def save_config(self):
        cfg = {
            "url": self.url_var.get().strip(),
            "method": self.method_var.get(),
            "interval": int(self.interval_var.get() or 60),
            "retry": int(self.retry_var.get() or 0),
            "timeout": int(self.timeout_var.get() or 10),
            "body": self.body_txt.get("1.0", tk.END).strip(),
        }
        with open(self._config_path(), "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        self._log("配置已保存")

    # ---------- 核心逻辑 ----------
    def _render(self, text):
        """动态变量渲染"""
        now = datetime.now()
        import random
        repl = {
            "{timestamp}": str(int(now.timestamp())),
            "{date}": now.strftime("%Y-%m-%d"),
            "{time}": now.strftime("%H:%M:%S"),
            "{rand}": str(random.randint(100000, 999999)),
        }
        for k, v in repl.items():
            text = text.replace(k, v)
        return text

    def _do_request(self):
        url = self._render(self.url_var.get().strip())
        method = self.method_var.get()
        body_raw = self._render(self.body_txt.get("1.0", tk.END).strip())
        timeout = int(self.timeout_var.get() or 10)
        headers = {"User-Agent": "TaskRunner/1.0"}
        if body_raw and method == "POST":
            try:
                json.loads(body_raw)
                headers["Content-Type"] = "application/json"
                data = body_raw
            except Exception:
                data = body_raw
        else:
            data = None
        resp = requests.request(method, url, headers=headers, data=data, timeout=timeout)
        return resp

    def _worker(self):
        interval = max(1, int(self.interval_var.get() or 60))
        max_retry = int(self.retry_var.get() or 0)
        while self.running:
            ok = False
            tries = 0
            while tries <= max_retry and self.running:
                try:
                    r = self._do_request()
                    ok = (200 <= r.status_code < 400)
                    self._log(f"[{datetime.now():%H:%M:%S}] HTTP {r.status_code} ({len(r.content)} bytes)")
                    break
                except Exception as e:
                    tries += 1
                    self.retry += 1
                    self._log(f"[{datetime.now():%H:%M:%S}] 失败: {str(e)[:60]} (重试 {tries}/{max_retry})")
                    time.sleep(min(2 * tries, 10))
            if ok:
                self.success += 1
            else:
                self.fail += 1
            self._update_count()
            # 倒计时等待（可中断）
            for _ in range(interval * 2):
                if not self.running:
                    return
                time.sleep(0.5)

    def test_once(self):
        try:
            r = self._do_request()
            self._log(f"[测试] HTTP {r.status_code} ({len(r.content)} bytes)")
            messagebox.showinfo("测试结果", f"HTTP {r.status_code}\n返回 {len(r.content)} 字节")
        except Exception as e:
            self._log(f"[测试] 失败: {str(e)[:80]}")
            messagebox.showerror("测试失败", str(e)[:200])

    # ---------- 控制 ----------
    def start(self):
        if not self.url_var.get().strip():
            messagebox.showwarning("提示", "请先填写请求地址")
            return
        self.running = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.status_var.set(f"状态：运行中 · 每 {self.interval_var.get()} 秒执行")
        self._log(f"[{datetime.now():%H:%M:%S}] 任务启动")
        self.thread = threading.Thread(target=self._worker, daemon=True)
        self.thread.start()

    def stop(self):
        self.running = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_var.set("状态：已停止")
        self._log(f"[{datetime.now():%H:%M:%S}] 任务停止")

    # ---------- 日志 ----------
    def _log(self, msg):
        self.log_txt.config(state=tk.NORMAL)
        self.log_txt.insert(tk.END, msg + "\n")
        self.log_txt.see(tk.END)
        self.log_txt.config(state=tk.DISABLED)
        self._file_log(msg)

    def _file_log(self, msg):
        try:
            os.makedirs(os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), LOG_DIR), exist_ok=True)
            fp = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), LOG_DIR, datetime.now().strftime("%Y-%m-%d.log"))
            with open(fp, "a", encoding="utf-8") as f:
                f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {msg}\n")
        except Exception:
            pass

    def open_logs(self):
        d = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), LOG_DIR)
        os.makedirs(d, exist_ok=True)
        os.startfile(d)

    def _update_count(self):
        self.count_var.set(f"成功 {self.success} | 失败 {self.fail}")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
