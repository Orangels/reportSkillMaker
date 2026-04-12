#!/usr/bin/env python3
"""
检查 section_manifest.json 中哪些 section 还未生成对应的 .py 文件。

用法：
    python3 check_sections.py <SESSION_DIR>

输出：
    ALL_DONE                                    — 所有 section 均已完成
    REMAINING=N, NEXT_BATCH=['id1', 'id2', ...]  — 还有 N 个未完成，本批处理前4个
"""
import json
import os
import sys

if len(sys.argv) != 2:
    print("Usage: python3 check_sections.py <SESSION_DIR>", file=sys.stderr)
    sys.exit(1)

session_dir = sys.argv[1]
manifest_path = os.path.join(session_dir, "section_manifest.json")

if not os.path.exists(manifest_path):
    print(f"ERROR: section_manifest.json not found in {session_dir}", file=sys.stderr)
    sys.exit(1)

with open(manifest_path, encoding="utf-8") as f:
    manifest = json.load(f)

remaining = [
    s["id"]
    for s in manifest["sections"]
    if not os.path.exists(os.path.join(session_dir, f"section_{s['id']}.py"))
]

if not remaining:
    print("ALL_DONE")
else:
    batch = remaining[:4]
    print(f"REMAINING={len(remaining)}, NEXT_BATCH={batch}")
