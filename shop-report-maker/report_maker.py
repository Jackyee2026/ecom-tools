# -*- coding: utf-8 -*-
"""
店铺报表自动化（SKU 3）v1.0
对应需求：猪八戒「Excel自动处理/进销存报表」高频品类（营口卖家样板同款赛道）
功能：
  1. 批量载入店铺订单明细表（xlsx/csv，多文件自动合并）
  2. 智能识别列（日期/商品/数量/金额——支持常见列名变体）
  3. 一键生成三张汇总表：按日汇总/按商品汇总/按月汇总
  4. 自动插入图表（日销售折线+商品TOP10柱状）
  5. 输出格式化 Excel 报表（带样式），原件不动
GUI：tkinter；打包：pyinstaller --onefile --noconsole
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import sys
from datetime import datetime

APP_TITLE = "店铺报表自动化 v1.0"

# 列名智能识别映射（常见变体 → 标准名）
COL_MAP = {
    "date": ["日期", "下单时间", "成交时间", "付款时间", "创建时间", "订单日期", "时间", "date"],
    "product": ["商品", "商品名称", "宝贝标题", "产品名称", "货品", "品名", "商品标题", "product", "item"],
    "qty": ["数量", "购买数量", "件数", "销量", "num", "qty", "quantity"],
    "amount": ["金额", "订单金额", "成交金额", "实付金额", "销售额", "总价", "amount", "price"],
}


def app_dir():
    return os.path.dirname(os.path.abspath(sys.argv[0]))


def norm(s):
    return str(s).strip().lower().replace(" ", "")


def detect_col(columns, kind, exclude=()):
    """在原始列名里找 kind 对应的标准列（exclude 避让已占用列）"""
    cands = [norm(c) for c in COL_MAP[kind]]
    pool = [c for c in columns if c not in exclude]
    for c in pool:
        if norm(c) in cands:
            return c
    # 模糊包含
    for c in pool:
        for cand in cands:
            if cand in norm(c) or norm(c) in cand:
                return c
    return None


STD = {"date": "_日期", "product": "_商品", "qty": "_数量", "amount": "_金额"}


class App:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("860x600")
        root.minsize(760, 540)
        self.files = []
        self.df = None
        self.colmap = {}
        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill=tk.X)
        ttk.Button(top, text="① 选择订单表(可多选)", command=self.pick_files).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="② 载入识别", command=self.load_files).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="③ 生成报表", command=self.gen_report).pack(side=tk.LEFT, padx=4)
        self.status = tk.StringVar(value="状态：未载入")
        ttk.Label(top, textvariable=self.status, font=("微软雅黑", 10, "bold")).pack(side=tk.RIGHT)

        # 列映射区
        mp = ttk.LabelFrame(self.root, text="列识别结果（自动识别，可手动修正下拉）", padding=8)
        mp.pack(fill=tk.X, padx=8, pady=4)
        self.combos = {}
        labels = {"date": "日期列", "product": "商品列", "qty": "数量列", "amount": "金额列"}
        for i, (k, lb) in enumerate(labels.items()):
            ttk.Label(mp, text=lb + "：").grid(row=i // 2, column=(i % 2) * 2, sticky="e", padx=(10, 2), pady=3)
            cb = ttk.Combobox(mp, width=24, state="readonly")
            cb.grid(row=i // 2, column=(i % 2) * 2 + 1, sticky="w", pady=3)
            self.combos[k] = cb

        # 说明区
        info = ttk.LabelFrame(self.root, text="报表内容", padding=8)
        info.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)
        tk.Label(info, justify=tk.LEFT, anchor="nw", font=("微软雅黑", 10), text=(
            "生成一份 Excel 报表，含 4 个工作表：\n\n"
            "  📊 总览        总销售额 / 总订单数 / 客单价 / 日均销售 / TOP商品\n"
            "  📅 按日汇总    每日销售额+订单数+件数，附折线图\n"
            "  📦 按商品汇总  每个商品的销量+销售额+占比，附TOP10柱状图\n"
            "  🗓 按月汇总    月度销售趋势对比\n\n"
            "适合：淘宝/拼多多/抖店/京东后台导出的订单明细表（自动识别常见列名）。\n"
            "原文件绝不修改，报表输出到 exe 同目录「报表输出」文件夹。"
        )).pack(anchor="nw")

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

    def pick_files(self):
        fs = filedialog.askopenfilenames(
            title="选择订单明细表（可多选）",
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
            allcols = []
            for fp in self.files:
                d = self._read_one(fp)
                # ⭐逐文件识别并归一化列名（不同平台导出列名不同，必须单文件识别再合并）
                used = set()
                ren = {}
                hit = {}
                for k in ("date", "product", "qty", "amount"):
                    c = detect_col(d.columns, k, exclude=used)
                    if c:
                        used.add(c)
                        ren[c] = STD[k]
                        hit[k] = c
                d = d.rename(columns=ren)
                dfs.append(d)
                allcols.extend(list(d.columns))
                self._log(f"载入 {os.path.basename(fp)}：{len(d)} 行 | 识别 {list(hit.keys())}")
            self.df = pd.concat(dfs, ignore_index=True) if len(dfs) > 1 else dfs[0]
            cols = list(dict.fromkeys(allcols))
            self.colmap = {}
            for k in ("date", "product", "qty", "amount"):
                present = STD[k] if STD[k] in self.df.columns else None
                self.colmap[k] = present
                self.combos[k]["values"] = [""] + cols
                self.combos[k].set(present or "")
            self._log(f"合并 {len(self.df)} 行；归一化列 {[v for v in self.colmap.values() if v]}")
            if not self.colmap.get("amount"):
                messagebox.showwarning("提醒", "未识别到金额列，请手动在下拉框指定")
            self.status.set(f"状态：已载入 {len(self.df)} 行")
        except Exception as e:
            messagebox.showerror("载入失败", str(e)[:300])
            self._log(f"载入失败: {str(e)[:100]}")

    def gen_report(self):
        if self.df is None:
            messagebox.showwarning("提示", "请先载入文件")
            return
        # 读用户最终选择
        cm = {k: (cb.get() or None) for k, cb in self.combos.items()}
        if not cm.get("amount"):
            messagebox.showwarning("提示", "金额列必须指定")
            return
        try:
            df = self.df.copy()
            df["_金额"] = pd.to_numeric(df[cm["amount"]], errors="coerce").fillna(0)
            if cm.get("qty"):
                df["_数量"] = pd.to_numeric(df[cm["qty"]], errors="coerce").fillna(0)
            else:
                df["_数量"] = 1
            if cm.get("date"):
                # ⭐pandas 3.0 严格模式：带时间的字符串会返回 NaT，必须 format="mixed"
                try:
                    df["_日期"] = pd.to_datetime(df[cm["date"]], errors="coerce", format="mixed")
                except Exception:
                    df["_日期"] = pd.to_datetime(df[cm["date"]], errors="coerce")
                bad = df["_日期"].isna().sum()
                if bad:
                    self._log(f"提示：{bad} 行日期无法解析已排除")
                df = df.dropna(subset=["_日期"])
                df["_日"] = df["_日期"].dt.date
                df["_月"] = df["_日期"].dt.strftime("%Y-%m")
            else:
                df["_日"] = "未知日期"
                df["_月"] = "未知月份"
            if cm.get("product"):
                df["_商品"] = df[cm["product"]].astype(str).str.strip()
            else:
                df["_商品"] = "全部商品"

            outdir = os.path.join(app_dir(), "报表输出")
            os.makedirs(outdir, exist_ok=True)
            fp = os.path.join(outdir, f"店铺报表_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
            build_report(df, fp)
            self._log(f"报表已生成：{fp}")
            self.status.set("状态：报表已生成")
            if messagebox.askyesno("生成成功", f"报表已保存到：\n{fp}\n\n打开所在文件夹？"):
                os.startfile(outdir)
        except PermissionError:
            messagebox.showerror("失败", "文件被占用，请关闭 Excel 后重试")
        except Exception as e:
            messagebox.showerror("生成失败", str(e)[:300])
            self._log(f"生成失败: {str(e)[:120]}")


def build_report(df, fp):
    """核心报表引擎（独立函数，便于无头测试）"""
    from openpyxl import Workbook
    from openpyxl.chart import LineChart, BarChart, Reference
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils.dataframe import dataframe_to_rows

    # 三张汇总
    daily = df.groupby("_日").agg(销售额=("_金额", "sum"), 订单数=("_金额", "count"), 件数=("_数量", "sum")).reset_index()
    daily.columns = ["日期", "销售额", "订单数", "件数"]
    daily = daily.sort_values("日期")
    prod = df.groupby("_商品").agg(销售额=("_金额", "sum"), 件数=("_数量", "sum"), 订单数=("_金额", "count")).reset_index()
    prod["销售额占比"] = (prod["销售额"] / prod["销售额"].sum() * 100).round(1).astype(str) + "%"
    prod = prod.sort_values("销售额", ascending=False)
    monthly = df.groupby("_月").agg(销售额=("_金额", "sum"), 订单数=("_金额", "count"), 件数=("_数量", "sum")).reset_index()
    monthly.columns = ["月份", "销售额", "订单数", "件数"]
    monthly = monthly.sort_values("月份")

    total_amt = df["_金额"].sum()
    total_ord = len(df)
    n_days = max(1, daily["日期"].nunique())
    top_prod = prod.iloc[0] if len(prod) else None

    wb = Workbook()
    hfill = PatternFill("solid", fgColor="2F5597")
    hfont = Font(color="FFFFFF", bold=True, size=11)
    tfont = Font(bold=True, size=14, color="2F5597")

    def put_df(ws, d, start=1):
        for j, col in enumerate(d.columns, 1):
            c = ws.cell(row=start, column=j, value=str(col))
            c.fill, c.font, c.alignment = hfill, hfont, Alignment(horizontal="center")
        for i, row in enumerate(dataframe_to_rows(d, index=False, header=False), start + 1):
            for j, v in enumerate(row, 1):
                ws.cell(row=i, column=j, value=v)
        for j, col in enumerate(d.columns, 1):
            ws.column_dimensions[ws.cell(row=start, column=j).column_letter].width = max(12, min(40, int(d[col].astype(str).str.len().max() if len(d) else 12) + 4))

    # Sheet1 总览
    ws = wb.active
    ws.title = "总览"
    ws["A1"] = "店铺销售报表总览"
    ws["A1"].font = Font(bold=True, size=16, color="2F5597")
    ws["A2"] = f"生成时间：{datetime.now():%Y-%m-%d %H:%M}"
    rows = [
        ("总销售额", round(total_amt, 2)),
        ("总订单数", total_ord),
        ("客单价", round(total_amt / total_ord, 2) if total_ord else 0),
        ("覆盖天数", n_days),
        ("日均销售额", round(total_amt / n_days, 2)),
        ("总件数", int(df["_数量"].sum())),
        ("TOP1 商品", f"{top_prod['_商品']}（{top_prod['销售额']:.0f} 元）" if top_prod is not None else "-"),
    ]
    for i, (k, v) in enumerate(rows, 4):
        ws.cell(row=i, column=1, value=k).font = Font(bold=True)
        ws.cell(row=i, column=2, value=v)
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 36

    # Sheet2 按日汇总+折线图
    ws2 = wb.create_sheet("按日汇总")
    put_df(ws2, daily)
    if len(daily) > 1:
        ch = LineChart()
        ch.title = "每日销售额趋势"
        ch.height, ch.width = 8, 18
        data = Reference(ws2, min_col=2, min_row=1, max_row=len(daily) + 1)
        cats = Reference(ws2, min_col=1, min_row=2, max_row=len(daily) + 1)
        ch.add_data(data, titles_from_data=True)
        ch.set_categories(cats)
        ws2.add_chart(ch, "F2")

    # Sheet3 按商品汇总+柱状图
    ws3 = wb.create_sheet("按商品汇总")
    put_df(ws3, prod)
    top10 = min(10, len(prod))
    if top10 > 0:
        ch2 = BarChart()
        ch2.title = "商品销售额 TOP10"
        ch2.height, ch2.width = 8, 18
        data = Reference(ws3, min_col=2, min_row=1, max_row=top10 + 1)
        cats = Reference(ws3, min_col=1, min_row=2, max_row=top10 + 1)
        ch2.add_data(data, titles_from_data=True)
        ch2.set_categories(cats)
        ws3.add_chart(ch2, "G2")

    # Sheet4 按月汇总
    ws4 = wb.create_sheet("按月汇总")
    put_df(ws4, monthly)

    wb.save(fp)
    return fp


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
