#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""内置方案的文字对比度自查（WCAG 相对亮度）。

纯标准库。用于验证 references/builtin-styles.md 的「对比度红线」；
改动任何方案色值后跑一遍，确认没有把文字压在对比度不足的底色上。

用法:
    python check_contrast.py              # 全量比对，输出表格
    python check_contrast.py --strict     # 有 ❌ 时退出码 1（供 CI / 批处理）

判定标准：正文(<24px) ≥4.5:1；大字(≥24px) ≥3.0:1。
"""

from __future__ import annotations

import sys

# (方案, 底色名称/hex, 前景名称/hex, 该前景的最低字号档: "body" 或 "large")
CHECKS = [
    # ── B1 医疗学术蓝 ──
    ("B1", "白底 #FFFFFF", "正文 #1A2230", "body"),
    ("B1", "白底 #FFFFFF", "副文本 #4A5568", "body"),
    ("B1", "白底 #FFFFFF", "辅助 #6B7789", "body"),
    ("B1", "白底 #FFFFFF", "旧辅助 #8B97A8（已弃用·禁）", "large"),
    ("B1", "浅底 #F7F9FC", "副文本 #4A5568", "body"),
    ("B1", "主蓝 #1E4FA8", "白字 #FFFFFF", "body"),
    ("B1", "藏蓝 #0E3F8C", "白字 #FFFFFF", "body"),
    ("B1", "中蓝 #3D7BD9", "白字 #FFFFFF", "large"),
    ("B1", "施警红 #D9534F", "白字 #FFFFFF", "large"),
    # ── B2 临床青绿 ──
    ("B2", "白底 #FFFFFF", "正文 #14201E", "body"),
    ("B2", "浅底 #F1FAF8", "正文 #14201E", "body"),
    ("B2", "主青绿 #0F766E", "白字 #FFFFFF", "body"),
    ("B2", "深青 #115E59", "白字 #FFFFFF", "body"),
    ("B2", "亮青 #14B8A6", "深字 #14201E", "body"),
    ("B2", "亮青 #14B8A6", "白字 #FFFFFF（禁）", "large"),
    ("B2", "施警红 #DC2626", "白字 #FFFFFF", "large"),
    # ── B3 深蓝权威 ──
    ("B3", "白底 #FFFFFF", "Action #061F32", "body"),
    ("B3", "白底 #FFFFFF", "脚注 #667085", "body"),
    ("B3", "Navy #061F32", "白字 #FFFFFF", "body"),
    ("B3", "皇家蓝 #001965", "白字 #FFFFFF", "body"),
    ("B3", "Cyan #08A6F6", "深字 #061F32", "body"),
    ("B3", "Cyan #08A6F6", "白字 #FFFFFF（禁）", "large"),
    ("B3", "浅蓝 #99BDED", "深字 #061F32", "body"),
    ("B3", "浅蓝 #99BDED", "白字 #FFFFFF（禁）", "large"),
    # ── B4 高级灰文献 ──
    ("B4", "白底 #FFFFFF", "正文 #000000", "body"),
    ("B4", "卡片 #F4F4F4", "正文 #000000", "body"),
    ("B4", "主色 #586874", "白字 #FFFFFF", "body"),
    ("B4", "深蓝 #1B3868", "白字 #FFFFFF", "body"),
    ("B4", "浅灰青 #AAC1D3", "黑字 #000000", "body"),
    ("B4", "浅灰青 #AAC1D3", "白字 #FFFFFF（禁）", "large"),
    ("B4", "中灰青 #8192A2", "白字 #FFFFFF", "large"),
    ("B4", "强调红 #C00000", "白字 #FFFFFF", "large"),
    # ── B5 红金政务 ──
    ("B5", "白底 #FFFFFF", "正文 #171616", "body"),
    ("B5", "主红 #E63F26", "白字 #FFFFFF", "large"),
    ("B5", "深朱红 #C42819", "白字 #FFFFFF", "body"),
    ("B5", "高光红 #EE5530", "白字 #FFFFFF", "large"),
    ("B5", "粉红橙 #F26B40", "深字 #171616", "body"),
    ("B5", "暗红 #A81810", "白字 #FFFFFF", "body"),
    ("B5", "主金 #E19825", "深字 #171616", "body"),
    ("B5", "主金 #E19825", "白字 #FFFFFF（禁）", "large"),
    ("B5", "浅金 #FFC000", "深字 #171616", "body"),
]


def _lin(c: int) -> float:
    c = c / 255
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4


def _lum(hex_color: str) -> float:
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return 0.2126 * _lin(r) + 0.7152 * _lin(g) + 0.0722 * _lin(b)


def ratio(fg: str, bg: str) -> float:
    a, b = _lum(fg), _lum(bg)
    hi, lo = max(a, b), min(a, b)
    return (hi + 0.05) / (lo + 0.05)


def hexof(label: str) -> str:
    """从 `名称 #RRGGBB` 中取色值。"""
    i = label.find("#")
    return label[i:i + 7]


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    strict = "--strict" in (argv if argv is not None else sys.argv[1:])

    print(f"{'方案':<5}{'底色':<26}{'前景':<34}{'对比度':>7}  {'判定':<12}{'要求'}")
    print("-" * 104)
    fails = 0
    cur = None
    for plan, bg_label, fg_label, need in CHECKS:
        if cur is not None and plan != cur:
            print()
        cur = plan
        r = ratio(hexof(fg_label), hexof(bg_label))
        lim = 4.5 if need == "body" else 3.0
        enough = r >= lim
        if "（禁）" in fg_label or "已弃用" in fg_label:
            # 反例项：对比度**不足**才是符合预期
            ok = not enough
            verdict = "OK(反例已避免)" if ok else "WARN(意外够清晰)"
        else:
            ok = enough
            verdict = "OK" if ok else "FAIL"
        if not ok:
            fails += 1
        print(f"{plan:<5}{bg_label:<26}{fg_label:<34}{r:>6.2f}  {verdict:<12}≥{lim}")

    print()
    print(f"合计 {len(CHECKS)} 项，异常 {fails} 项。")
    if fails == 0:
        print("✅ 五套方案的文字对比度全部符合红线。")
    else:
        print("❌ 存在不达标项，请调整 references/builtin-styles.md 中的色值。")
    return 1 if (strict and fails) else 0


if __name__ == "__main__":
    sys.exit(main())
