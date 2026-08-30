# 历史留档：每次排版的产物持久化到 out/history/<run_id>/，供演示界面回看与下载。
# 结构：out/history/<run_id>/{formatted.docx, report.md, meta.json}
# 公文场景话术：全程留痕、可追溯（审计友好）。

import json
import os
import shutil
from datetime import datetime

HISTORY_ROOT = os.path.join("out", "history")


def save_run(out_path, report_path, meta):
    """把本次排版产物拷入历史目录，写入 meta.json，返回 run_id。"""
    run_id = datetime.now().strftime("%Y%m%d-%H%M%S")
    root = os.path.join(HISTORY_ROOT, run_id)
    n = 1
    while os.path.exists(root):  # 同一秒内重复运行避免撞目录
        n += 1
        root = os.path.join(HISTORY_ROOT, "%s-%d" % (run_id, n))
        run_id = "%s-%d" % (run_id, n)
    os.makedirs(root)

    if os.path.exists(out_path):
        shutil.copy2(out_path, os.path.join(root, "formatted.docx"))
    if report_path and os.path.exists(report_path):
        shutil.copy2(report_path, os.path.join(root, "report.md"))

    record = dict(meta)
    record["run_id"] = run_id
    record["time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(os.path.join(root, "meta.json"), "w", encoding="utf-8") as f:
        json.dump(record, f, ensure_ascii=False, indent=2)
    return run_id


def list_runs():
    """按时间倒序返回历史记录：[{run_id, time, ..., docx, report}]。"""
    if not os.path.isdir(HISTORY_ROOT):
        return []
    runs = []
    for d in os.listdir(HISTORY_ROOT):
        meta_path = os.path.join(HISTORY_ROOT, d, "meta.json")
        if not os.path.isfile(meta_path):
            continue
        try:
            with open(meta_path, encoding="utf-8") as f:
                meta = json.load(f)
        except (OSError, ValueError):
            continue
        meta["docx"] = os.path.join(HISTORY_ROOT, d, "formatted.docx")
        meta["report"] = os.path.join(HISTORY_ROOT, d, "report.md")
        if not os.path.isfile(meta["docx"]):
            continue  # 产物已丢失的记录不展示
        runs.append(meta)
    runs.sort(key=lambda m: m.get("time", ""), reverse=True)
    return runs
