#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""模板检测：判断本次该用哪套 PPT 模板，并给出该问的问题与选项。

纯标准库。由 wechat-to-medical-ppt 技能的 Step 0.5 调用。

用法:
    python check_template.py [工作空间根目录]        # 缺省用当前目录
    python check_template.py --list-builtins        # 只打印内置方案菜单

输出（stdout，UTF-8）：检测结论 + 建议的提问分支
退出码始终为 0（检测失败时按"无模板"处理并给出提示）。
"""

from __future__ import annotations

import os
import re
import sys

TPL_DIR = "04_模板"
CONFIG = "_模板配置.md"
EXTS = (".pptx", ".potx")
LEGACY = (".ppt", ".pot")

# 内置方案库，与 references/builtin-styles.md 保持一致
BUILTINS = [
    ("B1", "医疗学术蓝", "指南/共识解读、通用学术讲课（默认）"),
    ("B2", "临床青绿", "疾病科普、患者教育、护理感控、生活方式"),
    ("B3", "深蓝权威", "制度流程宣贯、质控通报、科室管理汇报"),
    ("B4", "高级灰文献", "研究生组会、文献评读、循证证据"),
    ("B5", "红金政务", "党建、思想政治、行政表彰"),
]
DEFAULT_BUILTIN = "B1"


def builtin_label(code: str, with_desc: bool = False) -> str:
    for c, n, d in BUILTINS:
        if c == code:
            return f"内置 · {c} {n}" + (f"（{d}）" if with_desc else "")
    return f"内置 · {DEFAULT_BUILTIN} {BUILTINS[0][1]}"


def short_name(code: str) -> str:
    for c, n, _ in BUILTINS:
        if c == code:
            return n
    return BUILTINS[0][1]


def parse_builtin(value: str) -> str | None:
    """从配置里的「默认模板」取值中解析出内置方案代号；非内置返回 None。"""
    if not value or not value.startswith("内置"):
        return None
    m = re.search(r"\bB([1-5])\b", value)
    if m:
        return f"B{m.group(1)}"
    for c, n, _ in BUILTINS:
        if n in value:
            return c
    return DEFAULT_BUILTIN


def read_config(path: str) -> dict:
    cfg = {"default": "内置", "file": "—"}
    if not os.path.isfile(path):
        return cfg
    try:
        text = open(path, encoding="utf-8").read()
    except OSError:
        return cfg
    m = re.search(r"\*\*默认模板\*\*[：:]\s*(.+)", text)
    if m:
        cfg["default"] = m.group(1).strip()
    m = re.search(r"\*\*默认模板文件\*\*[：:]\s*(.+)", text)
    if m:
        cfg["file"] = m.group(1).strip()
    return cfg


def scan(tpl_dir: str) -> tuple[list[str], list[str]]:
    """返回 (可用的 pptx/potx 列表, 需要转换的旧格式列表)。"""
    ok, legacy = [], []
    if not os.path.isdir(tpl_dir):
        return ok, legacy
    for name in sorted(os.listdir(tpl_dir)):
        p = os.path.join(tpl_dir, name)
        if not os.path.isfile(p) or name.startswith("~$"):
            continue
        low = name.lower()
        if low.endswith(EXTS):
            ok.append(name)
        elif low.endswith(LEGACY):
            legacy.append(name)
    return ok, legacy


def print_builtins() -> None:
    print("# 内置方案库（无需文件，直接选用）")
    for c, n, d in BUILTINS:
        mark = "  ← 默认" if c == DEFAULT_BUILTIN else ""
        print(f"  {c}  {n:<8} {d}{mark}")
    print("\n规范详见：<技能目录>/references/builtin-styles.md")


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    args = argv if argv is not None else sys.argv[1:]

    if "--list-builtins" in args:
        print_builtins()
        return 0
    args = [a for a in args if not a.startswith("--")]
    root = os.path.abspath((args or ["."])[0])

    tpl_dir = os.path.join(root, TPL_DIR)
    cfg_path = os.path.join(tpl_dir, CONFIG)

    config = read_config(cfg_path)
    found, legacy = scan(tpl_dir)

    code = parse_builtin(config["default"])          # 非 None = 默认是内置方案
    default_is_builtin = code is not None
    cur_builtin = builtin_label(code if code else DEFAULT_BUILTIN)

    cfg_file = config["file"] if config["file"] not in ("—", "") else None
    cfg_exists = bool(cfg_file) and os.path.isfile(
        cfg_file if os.path.isabs(cfg_file) else os.path.join(root, cfg_file)
    )

    out = []
    A = out.append
    A("# 模板检测结果")
    A(f"- 工作空间：`{root}`")
    A(f"- 配置里的默认模板：**{config['default']}**"
      + (f"（文件：`{config['file']}`）" if cfg_file else "")
      + (f" → 解析为 {cur_builtin}" if default_is_builtin else " → 用户自有模板"))
    A(f"- `{TPL_DIR}/` 下的可用模板（{len(found)}）：" +
      ("、".join(f"`{f}`" for f in found) if found else "无"))
    if legacy:
        A(f"- ⚠️ 旧格式需先另存为 .pptx：{'、'.join(f'`{f}`' for f in legacy)}")
    A("")
    A("## 内置方案库（本次可选）")
    for c, n, d in BUILTINS:
        A(f"  - {c} {n} — {d}")
    A("")

    # 分支判定
    if found and not default_is_builtin and cfg_exists:
        branch = "A"
        question = "本次课件用哪套模板？"
        options = [
            f"用刚放入的 {found[0]}",
            f"沿用上次的 {config['default']}",
            f"{builtin_label(DEFAULT_BUILTIN)}（默认）",
            "换别的内置方案（B2 青绿 / B3 深蓝 / B4 高级灰 / B5 红金）",
        ]
    elif found:
        branch = "B"
        question = f"本次用你的模板「{found[0]}」吗？"
        options = [
            f"用 {found[0]}",
            f"{builtin_label(DEFAULT_BUILTIN)}（默认）",
            "换别的内置方案（B2 / B3 / B4 / B5）",
        ]
    elif not default_is_builtin and cfg_exists:
        branch = "C"
        question = f"没有新模板，本次沿用上次的「{config['default']}」吗？"
        options = [
            f"沿用 {config['default']}",
            f"{builtin_label(DEFAULT_BUILTIN)}（默认）",
            "换别的内置方案（B2 / B3 / B4 / B5）",
        ]
    elif not default_is_builtin and not cfg_exists:
        branch = "E"
        question = "上次用的模板文件没找到，本次用哪套？"
        options = [
            f"{builtin_label(DEFAULT_BUILTIN)}（默认）",
            "换别的内置方案（B2 / B3 / B4 / B5）",
            "我重新放一份到 04_模板/",
        ]
    else:
        branch = "D"
        question = f"本次用哪套内置方案？（默认 {builtin_label(DEFAULT_BUILTIN)}）"
        options = [
            f"{builtin_label(DEFAULT_BUILTIN)}（推荐：学术讲课通用）",
            f"内置 · B2 临床青绿（科普/患者教育）",
            f"内置 · B3 深蓝权威（制度/质控宣贯）",
            "其他（B4 高级灰文献 / B5 红金政务 / 我放自己的模板）",
        ]

    A(f"## 分支 {branch}")
    A(f"提问：**{question}**")
    A("选项（用 AskUserQuestion 呈现，单选，最多 4 项）：")
    for i, o in enumerate(options, 1):
        A(f"  {i}. {o}")
    A("")
    A("确认后请：① 更新 `04_模板/_模板配置.md` 的「默认模板/默认模板文件/最近一次使用」并追加历史一行"
      "（内置方案写成 `内置 · B1 医疗学术蓝`）；"
      "② 若选用户模板，跑 `scripts/extract_template_design.py` 生成设计契约；"
      "③ 若选内置方案，打开 `references/builtin-styles.md` 读该方案规范；"
      "④ 在汇报里写明「本次使用模板：X」。")

    print("\n".join(out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
