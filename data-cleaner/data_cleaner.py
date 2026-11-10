# -*- coding: utf-8 -*-
"""
数据清洗助手（SKU 2）v1.0
对应需求：猪八戒「数据处理软件·批量清洗转换」（高频品类）
功能：
  1. Excel(.xlsx)/CSV 批量载入（可多选文件）
  2. 去重（全行/指定列）
  3. 空值处理（删行/填充指定值）
  4. 去首尾空格（全表文本列）
  5. 查找替换（指定列或全表）
  6. 多文件纵向合并（表头一致）
  7. 清洗结果输出新文件（绝不覆盖原件）
GUI：tkinter；打包：pyinstaller --onefile --noconsole
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import sys
from datetime import datetime

APP_TITLE = "数据清洗助手 v1.0"


def app_dir():
    return os.path.dirname(os.path.abspath(sys.argv[0]))


class App:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("980x700")
        root.minsize(880, 620)
        self.files = []
        self.df = None          # 当前合并后的工作表
        self._build_ui()

    # ---------- UI ----------
    def _build_ui(self):
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill=tk.X)

        ttk.Button(top, text="① 选择文件(Excel/CSV)", command=self.pick_files).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="② 载入并预览", command=self.load_files).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="③ 执行清洗", command=self.run_clean).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="④ 导出结果", command=self.export).pack(side=tk.LEFT, padx=4)
        self.status = tk.StringVar(value="状态：未载入文件")
        ttk.Label(top, textvariable=self.status, font=("微软雅黑", 10, "bold")).pack(side=tk.RIGHT)

        # 清洗选项
        opt = ttk.LabelFrame(self.root, text="清洗选项（勾选后按③执行）", padding=8)
        opt.pack(fill=tk.X, padx=8, pady=4)

        # 行1
        r1 = ttk.Frame(opt); r1.pack(fill=tk.X, pady=2)
        self.v_dedup = tk.BooleanVar()
        ttk.Checkbutton(r1, text="去重", variable=self.v_dedup).pack(side=tk.LEFT)
        ttk.Label(r1, text="依据列：").pack(side=tk.LEFT, padx=(8, 2))
        self.dedup_col = ttk.Combobox(r1, width=16, state="readonly")
        self.dedup_col.pack(side=tk.LEFT)
        ttk.Label(r1, text="（留空=整行去重）", foreground="#888").pack(side=tk.LEFT, padx=4)

        self.v_space = tk.BooleanVar()
        ttk.Checkbutton(r1, text="去除文本首尾空格", variable=self.v_space).pack(side=tk.LEFT, padx=(24, 0))

        # 行2
        r2 = ttk.Frame(opt); r2.pack(fill=tk.X, pady=2)
        self.v_null = tk.BooleanVar()
        ttk.Checkbutton(r2, text="空值处理：", variable=self.v_null).pack(side=tk.LEFT)
        self.null_mode = ttk.Combobox(r2, values=["删除含空值的行", "填充为指定值"], width=16, state="readonly")
        self.null_mode.current(0)
        self.null_mode.pack(side=tk.LEFT, padx=4)
        ttk.Label(r2, text="填充值：").pack(side=tk.LEFT, padx=(8, 2))
        self.null_fill = ttk.Entry(r2, width=12)
        self.null_fill.pack(side=tk.LEFT)

        # 行3
        r3 = ttk.Frame(opt); r3.pack(fill=tk.X, pady=2)
        self.v_replace = tk.BooleanVar()
        ttk.Checkbutton(r3, text="查找替换：", variable=self.v_replace).pack(side=tk.LEFT)
        ttk.Label(r3, text="查找").pack(side=tk.LEFT, padx=(8, 2))
        self.rep_find = ttk.Entry(r3, width=14)
        self.rep_find.pack(side=tk.LEFT)
        ttk.Label(r3, text="替换为").pack(side=tk.LEFT, padx=4)
        self.rep_to = ttk.Entry(r3, width=14)
        self.rep_to.pack(side=tk.LEFT)
        ttk.Label(r3, text="（全表文本生效）", foreground="#888").pack(side=tk.LEFT, padx=4)

        # 预览表
        pv = ttk.LabelFrame(self.root, text="数据预览（前50行）", padding=4)
        pv.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)
        self.tree = ttk.Treeview(pv, show="headings")
        vsb = ttk.Scrollbar(pv, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(pv, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        pv.rowconfigure(0, weight=1); pv.columnconfigure(0, weight=1)

        # 日志
        lf = ttk.LabelFrame(self.root, text="操作日志", padding=4)
        lf.pack(fill=tk.X, padx=8, pady=(0, 8))
        self.log_txt = tk.Text(lf, height=6, state=tk.DISABLED, font=("Consolas", 9))
        self.log_txt.pack(fill=tk.X)

    def _log(self, m):
        line = f"[{datetime.now():%H:%M:%S}] {m}"
        self.log_txt.config(state=tk.NORMAL)
        self.log_txt.insert(tk.END, line + "\n")
        self.log_txt.see(tk.END)
        self.log_txt.config(state=tk.DISABLED)
        print(line)

    # ---------- 载入 ----------
    def pick_files(self):
        fs = filedialog.askopenfilenames(
            title="选择 Excel/CSV 文件（可多选）",
            filetypes=[("数据文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")])
        if fs:
            self.files = list(fs)
            self.status.set(f"状态：已选 {len(self.files)} 个文件")
            self._log(f"已选择 {len(self.files)} 个文件")

    def _read_one(self, fp):
        ext = os.path.splitext(fp)[1].lower()
        if ext == ".csv":
            for enc in ("utf-8-sig", "gbk", "utf-8"):
                try:
                    return pd.read_csv(fp, encoding=enc)
                except UnicodeDecodeError:
                    continue
            return pd.read_csv(fp, encoding="utf-8", errors="replace")
        return pd.read_excel(fp)

    def load_files(self):
        if not self.files:
            messagebox.showwarning("提示", "请先选择文件")
            return
        try:
            dfs = []
            for fp in self.files:
                d = self._read_one(fp)
                d["__来源文件__"] = os.path.basename(fp)
                dfs.append(d)
                self._log(f"载入 {os.path.basename(fp)}：{len(d)} 行 × {len(d.columns)-1} 列")
            if len(dfs) == 1:
                self.df = dfs[0]
            else:
                self.df = pd.concat(dfs, ignore_index=True)
                self._log(f"已纵向合并 {len(dfs)} 个文件 → {len(self.df)} 行")
            self._refresh_preview()
            self.dedup_col["values"] = [""] + [c for c in self.df.columns if c != "__来源文件__"]
            self.status.set(f"状态：已载入 {len(self.df)} 行 × {len(self.df.columns)} 列")
        except Exception as e:
            messagebox.showerror("载入失败", str(e)[:300])
            self._log(f"载入失败: {str(e)[:100]}")

    def _refresh_preview(self):
        cols = list(self.df.columns)[:30]
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=str(c))
            self.tree.column(c, width=110, anchor="w")
        for _, row in self.df.head(50).iterrows():
            self.tree.insert("", tk.END, values=[str(row[c])[:60] for c in cols])

    # ---------- 清洗 ----------
    def run_clean(self):
        if self.df is None:
            messagebox.showwarning("提示", "请先载入文件")
            return
        n0 = len(self.df)
        df = self.df.copy()
        acts = []

        if self.v_space.get():
            for c in df.columns:
                df[c] = df[c].map(lambda x: x.strip() if isinstance(x, str) else x)
            acts.append("去首尾空格")

        if self.v_replace.get() and self.rep_find.get():
            f, t = self.rep_find.get(), self.rep_to.get()
            for c in df.columns:
                df[c] = df[c].map(lambda x, f=f, t=t: x.replace(f, t) if isinstance(x, str) else x)
            acts.append(f"替换[{f}→{t}]")

        if self.v_null.get():
            if self.null_mode.get() == "删除含空值的行":
                before = len(df)
                df = df.dropna()
                acts.append(f"删空值行(-{before-len(df)})")
            else:
                df = df.fillna(self.null_fill.get())
                acts.append("空值填充")

        if self.v_dedup.get():
            before = len(df)
            col = self.dedup_col.get()
            if col:
                df = df.drop_duplicates(subset=[col])
                acts.append(f"按[{col}]去重(-{before-len(df)})")
            else:
                df = df.drop_duplicates()
                acts.append(f"整行去重(-{before-len(df)})")

        self.df = df.reset_index(drop=True)
        self._refresh_preview()
        self._log(f"清洗完成：{n0} → {len(df)} 行 ｜ 动作：{'、'.join(acts) if acts else '无（未勾选任何项）'}")
        self.status.set(f"状态：清洗完成 {len(df)} 行")

    # ---------- 导出 ----------
    def export(self):
        if self.df is None:
            messagebox.showwarning("提示", "没有可导出的数据")
            return
        outdir = os.path.join(app_dir(), "清洗结果")
        os.makedirs(outdir, exist_ok=True)
        fp = os.path.join(outdir, f"清洗结果_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        try:
            self.df.to_excel(fp, index=False)
            self._log(f"已导出：{fp}")
            self.status.set("状态：已导出")
            if messagebox.askyesno("导出成功", f"已保存到：\n{fp}\n\n打开所在文件夹？"):
                os.startfile(outdir)
        except PermissionError:
            messagebox.showerror("导出失败", "文件被占用，请关闭 Excel 后重试")
        except Exception as e:
            messagebox.showerror("导出失败", str(e)[:200])


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
