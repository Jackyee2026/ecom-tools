# -*- coding: utf-8 -*-
"""搭 ecom-tools 开源仓库结构"""
import os, shutil

BASE = r"E:\工具货架"
REPO = os.path.join(BASE, "ecom-tools")
SKUS = [
    ("SKU1-定时任务执行器", "task_runner.py", "timer-task-runner"),
    ("SKU2-数据清洗助手", "data_cleaner.py", "data-cleaner"),
    ("SKU3-店铺报表自动化", "report_maker.py", "shop-report-maker"),
]

os.makedirs(REPO, exist_ok=True)
for sk, main_py, en in SKUS:
    dst = os.path.join(REPO, en)
    os.makedirs(dst, exist_ok=True)
    # 主源码
    shutil.copy2(os.path.join(BASE, sk, main_py), os.path.join(dst, main_py))
    # 使用说明
    shutil.copy2(os.path.join(BASE, sk, "dist", "使用说明.txt"), os.path.join(dst, "使用说明.txt"))
    # spec 打包配置
    spec = [f for f in os.listdir(os.path.join(BASE, sk)) if f.endswith(".spec")]
    if spec:
        shutil.copy2(os.path.join(BASE, sk, spec[0]), os.path.join(dst, spec[0]))
    print("OK:", en)

print("repo ready:", REPO)
