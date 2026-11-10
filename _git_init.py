# -*- coding: utf-8 -*-
"""git init + 首次提交"""
import subprocess

GIT = r"C:\Program Files\Git\bin\git.exe"
REPO = r"E:\工具货架\ecom-tools"

def run(*args):
    r = subprocess.run([GIT] + list(args), cwd=REPO, capture_output=True, text=True, errors="ignore")
    return r.stdout + r.stderr

print(run("init"))
print(run("add", "-A"))
r = run("-c", "user.name=yifan-dev", "-c", "user.email=301052088@qq.com",
        "commit", "-m", "feat: first 3 e-commerce seller tools (timer task runner / data cleaner / shop report maker)")
print(r[:400])
print("LOG:", run("log", "--oneline"))
