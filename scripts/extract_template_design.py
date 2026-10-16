#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从用户自有 PPT 模板中抽取"视觉设计契约"（DESIGN.md 骨架 + 测量报告）。

纯标准库实现（zipfile + xml.etree），不依赖 LibreOffice / poppler / python-pptx，
因此在任意 Windows 机器上都能跑。

用法:
    python extract_template_design.py "<模板.pptx 或含模板的文件夹>" [-o <输出目录>]

产出（默认写到 <模板同级>/_提取/ 或 -o 指定目录）:
    TEMPLATE-MEASURE.md   测量报告：可直接用的精确值（🟢）+ 需人工/AI 判断的项（🟡🔴）
    DESIGN.md             设计契约骨架：测量字段已填，推断字段留 TODO
    assets/               从模板中抽出的候选固定资产（logo / 背景 / 装饰图）

坐标说明: 输出几何一律换算为 DSL 画布像素（EMU/9525）。
tencent-pptx 的 .slide 画布固定 1280x720 px，16:9 模板换算后可直接套用。
"""

from __future__ import annotations

import argparse
import os
import re
import struct
import sys
import zipfile
from collections import Counter
from xml.etree import ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

EMU_PER_PX = 9525          # OOXML: 1 px (96dpi) = 9525 EMU
EMU_PER_PT = 12700
DSL_W, DSL_H = 1280, 720   # tencent-pptx .slide 画布


def px(v) -> float:
    try:
        return round(float(v) / EMU_PER_PX, 1)
    except (TypeError, ValueError):
        return 0.0


def read_xml(zf: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(name))
    except Exception:
        return None


# ---------------------------------------------------------------- 图片尺寸

def image_size(data: bytes) -> tuple[int, int] | None:
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        w, h = struct.unpack(">II", data[16:24])
        return w, h
    if data[:3] == b"\xff\xd8\xff":
        i = 2
        while i + 9 < len(data):
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB):
                h, w = struct.unpack(">HH", data[i + 5:i + 9])
                return w, h
            if marker in (0xD8, 0xD9) or 0xD0 <= marker <= 0xD7:
                i += 2
                continue
            seg = struct.unpack(">H", data[i + 2:i + 4])[0]
            i += 2 + seg
        return None
    if data[:6] in (b"GIF87a", b"GIF89a"):
        w, h = struct.unpack("<HH", data[6:10])
        return w, h
    if data[:2] == b"BM":
        w, h = struct.unpack("<ii", data[18:26])
        return abs(w), abs(h)
    return None


def has_alpha_png(data: bytes) -> bool:
    """粗略判断 PNG 是否带 alpha（color type 4/6）。"""
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        return False
    return data[25] in (4, 6)


# ---------------------------------------------------------------- 主体抽取

def extract(tpl: str, outdir: str) -> dict:
    zf = zipfile.ZipFile(tpl)
    names = set(zf.namelist())
    res: dict = {"file": tpl}

    # ---- 画布
    pres = read_xml(zf, "ppt/presentation.xml")
    cx = cy = 0
    if pres is not None:
        sz = pres.find("p:sldSz", NS)
        if sz is not None:
            cx, cy = int(sz.get("cx", 0)), int(sz.get("cy", 0))
    res["emu"] = (cx, cy)
    res["px"] = (px(cx), px(cy))
    res["ratio"] = round(cx / cy, 4) if cy else 0
    res["scale_x"] = round(DSL_W / res["px"][0], 4) if res["px"][0] else 1
    res["scale_y"] = round(DSL_H / res["px"][1], 4) if res["px"][1] else 1

    slides = sorted(n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n))
    layouts = sorted(n for n in names if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n))
    masters = sorted(n for n in names if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n))
    res["counts"] = {"slides": len(slides), "layouts": len(layouts), "masters": len(masters)}

    # ---- 主题：配色 + 字体
    theme = None
    for cand in ("ppt/theme/theme1.xml",) + tuple(sorted(n for n in names if n.startswith("ppt/theme/theme"))):
        theme = read_xml(zf, cand)
        if theme is not None:
            res["theme_part"] = cand
            break
    scheme, fonts = {}, {}
    if theme is not None:
        cs = theme.find(".//a:clrScheme", NS)
        if cs is not None:
            for child in cs:
                tag = child.tag.split("}")[1]
                c = child.find("a:srgbClr", NS)
                if c is None:
                    c = child.find("a:sysClr", NS)
                    if c is not None:
                        scheme[tag] = "#" + (c.get("lastClr") or "000000").upper()
                    continue
                scheme[tag] = "#" + (c.get("val") or "").upper()
        fs = theme.find(".//a:fontScheme", NS)
        if fs is not None:
            for kind in ("majorFont", "minorFont"):
                node = fs.find(f"a:{kind}", NS)
                if node is None:
                    continue
                entry = {}
                for sub in ("latin", "ea", "cs"):
                    s = node.find(f"a:{sub}", NS)
                    if s is not None and s.get("typeface"):
                        entry[sub] = s.get("typeface")
                fonts[kind] = entry
    res["color_scheme"] = scheme
    res["font_scheme"] = fonts

    # ---- 统计色值 / 字体 / 字号（slide + layout + master 全量）
    color_c = Counter()
    latin_c = Counter()
    ea_c = Counter()
    size_c = Counter()
    bold_c = Counter()
    body_parts = slides + layouts + masters
    for part in body_parts:
        raw = zf.read(part).decode("utf-8", "ignore")
        for m in re.finditer(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', raw):
            color_c["#" + m.group(1).upper()] += 1
        for m in re.finditer(r'<a:sysClr[^>]*lastClr="([0-9A-Fa-f]{6})"', raw):
            color_c["#" + m.group(1).upper()] += 1
        for m in re.finditer(r'<a:latin typeface="([^"]+)"', raw):
            if m.group(1):
                latin_c[m.group(1)] += 1
        for m in re.finditer(r'<a:ea typeface="([^"]+)"', raw):
            if m.group(1):
                ea_c[m.group(1)] += 1
        for m in re.finditer(r'<a:(?:rPr|defRPr|endParaRPr)[^>]*\bsz="(\d+)"', raw):
            size_c[int(m.group(1))] += 1
        for m in re.finditer(r'<a:(?:rPr|defRPr)[^>]*\bb="1"', raw):
            bold_c[part] += 1
    res["colors"] = color_c.most_common(24)
    res["latin"] = latin_c.most_common(10)
    res["ea"] = ea_c.most_common(10)
    res["sizes"] = size_c.most_common(20)

    # 主题是否为 Office 出厂默认（此时主题色不代表设计意图，须改用高频实色）
    STOCK_THEME = {"#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5",
                   "#70AD47", "#44546A", "#E7E6E6", "#0563C1", "#954F72"}
    res["stock_theme"] = scheme.get("accent1", "").upper() in STOCK_THEME

    # 有效用色：按"有彩色 / 暗色 / 浅中性"分类，用于在主题不可信时反推真实配色
    def _rgb(c: str):
        try:
            return int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
        except ValueError:
            return None

    def _useful(c: str) -> bool:
        r = _rgb(c)
        if r is None:
            return False
        return not (min(r) > 232) and not (max(r) < 40)

    res["useful_colors"] = [c for c, _ in res["colors"] if _useful(c)]
    res["chromatic"] = [c for c, _ in res["colors"]
                        if _rgb(c) and (max(_rgb(c)) - min(_rgb(c))) >= 40
                        and min(_rgb(c)) > 8]
    res["dark_ink"] = [c for c, _ in res["colors"]
                       if _rgb(c) and max(_rgb(c)) < 90 and c != "#000000"]
    res["light_surface"] = [c for c, _ in res["colors"]
                            if _rgb(c) and min(_rgb(c)) > 232 and c != "#FFFFFF"]

    # 字号：加权中位数定 body，出现 ≥3 次的最小值定 caption
    valid = [(s, n) for s, n in size_c.most_common() if n >= 3] or size_c.most_common()
    valid.sort(key=lambda x: x[0])
    total = sum(n for _, n in valid) or 1
    acc, med = 0, valid[len(valid) // 2][0]
    for s, n in valid:
        acc += n
        if acc >= total / 2:
            med = s
            break
    res["size_roles"] = {
        "display": valid[-1][0] / 100,
        "body": med / 100,
        "caption": valid[0][0] / 100,
    }

    # ---- 母版 / 版式：背景、固定元素（logo、标题位、页脚、页码）
    fixed = {}
    for label, part in [("master", p) for p in masters] + [("layout", p) for p in layouts[:6]]:
        root = read_xml(zf, part)
        if root is None:
            continue
        info = {"bg": None, "pics": [], "ph": []}
        bg = root.find(".//p:cSld/p:bg", NS)
        if bg is not None:
            sc = bg.find(".//a:srgbClr", NS)
            info["bg"] = ("#" + sc.get("val").upper()) if sc is not None else "非纯色(图片/渐变)"
        for pic in root.iter("{%s}pic" % NS["p"]):
            xfrm = pic.find(".//a:xfrm", NS)
            off = xfrm.find("a:off", NS) if xfrm is not None else None
            ext = xfrm.find("a:ext", NS) if xfrm is not None else None
            if off is None or ext is None:
                continue
            info["pics"].append({
                "x": px(off.get("x")), "y": px(off.get("y")),
                "w": px(ext.get("cx")), "h": px(ext.get("cy")),
            })
        for ph in root.iter("{%s}ph" % NS["p"]):
            info["ph"].append(ph.get("type", "body"))
        fixed[part.split("/")[-1]] = info
    res["fixed"] = fixed

    # ---- 媒体资源（logo / 背景候选）
    media = []
    for n in sorted(names):
        if "/media/" not in n or n.endswith("/"):
            continue
        data = zf.read(n)
        wh = image_size(data)
        media.append({
            "part": n, "bytes": len(data),
            "px": f"{wh[0]}x{wh[1]}" if wh else "?",
            "alpha": has_alpha_png(data),
            "ext": os.path.splitext(n)[1].lower(),
        })
    res["media"] = media

    # ---- 抽固定资产
    asset_dir = os.path.join(outdir, "assets")
    os.makedirs(asset_dir, exist_ok=True)
    saved = []
    for m in media:
        if m["ext"] not in (".png", ".jpg", ".jpeg", ".gif", ".bmp"):
            continue
        keep = False
        maxdim = 0
        if m["px"] != "?":
            maxdim = max(int(v) for v in m["px"].split("x"))
        if m["ext"] in (".png", ".gif") and m["alpha"] and m["bytes"] < 900_000:
            keep = True                      # 透明图多为 logo / 图标
        if maxdim >= 1600:
            keep = True                      # 大图多为背景 / 主视觉
        if m["bytes"] >= 60_000:
            keep = True                      # 体积够大，值得留
        if 0 < maxdim <= 600:
            keep = True                      # 小尺寸：logo 候选（含不透明）
        if not keep or len(saved) >= 20:
            continue
        base = "asset_%03d_%s" % (len(saved) + 1, os.path.basename(m["part"]))
        with open(os.path.join(asset_dir, base), "wb") as fh:
            fh.write(zf.read(m["part"]))
        saved.append(base)
    res["assets_saved"] = saved

    zf.close()
    return res


# ---------------------------------------------------------------- 报告输出

def write_measure_md(res: dict, outdir: str) -> str:
    px_w, px_h = res["px"]
    L = []
    A = L.append
    A("# 模板测量报告 · TEMPLATE-MEASURE.md\n")
    A(f"- 源文件：`{os.path.basename(res['file'])}`")
    A(f"- 画布：**{px_w:g} x {px_h:g} px**（OOXML {res['emu'][0]} x {res['emu'][1]} EMU，"
      f"比例 {res['ratio']}）")
    A(f"- DSL 画布固定 1280x720 px → 横向缩放 {res['scale_x']}，纵向缩放 {res['scale_y']}")
    if abs(res["scale_x"] - 1) > 0.01 or abs(res["scale_y"] - 1) > 0.01:
        A(f"- ⚠️ 模板不是 16:9，所有坐标需乘以上述比例后再落到 .slide")
    A(f"- 页数 {res['counts']['slides']}，版式 {res['counts']['layouts']}，母版 {res['counts']['masters']}\n")

    A("## 🟢 精确值（取自 XML，可直接写入 DESIGN.md）\n")
    A("### 主题配色\n")
    A("| 角色 | 色值 |\n|---|---|")
    for k, v in res["color_scheme"].items():
        A(f"| {k} | `{v}` |")
    A("")
    if res.get("stock_theme"):
        A("> ⚠️ 主题色与 Office 出厂默认一致，说明该模板的真实配色写在页面上而非主题里 —— "
          "\n> 请以「实际用色频次」为准（DESIGN.md 已自动改用高频色）。\n")
    A("### 主题字体\n")
    for kind, entry in res["font_scheme"].items():
        A(f"- **{kind}**：" + "，".join(f"{k}={v}" for k, v in entry.items() if v))
    A("")
    A("### 实际用色频次（前 24）\n")
    A("| 色值 | 出现次数 |\n|---|---|")
    for c, n in res["colors"]:
        A(f"| `{c}` | {n} |")
    A("")
    A("### 字体使用频次\n")
    if res["latin"]:
        A("- 西文（latin）：" + "，".join(f"{f}({n})" for f, n in res["latin"]))
    if res["ea"]:
        A("- 中日韩（ea）：" + "，".join(f"{f}({n})" for f, n in res["ea"]))
    A("")
    A("### 字号分布（sz/100 = pt）\n")
    A("| pt | 次数 | 建议角色 |\n|---|---|---|")
    roles = res.get("size_roles", {})
    rmap = {round(roles.get("display", 0) * 100): "display（最大字）",
            round(roles.get("body", 0) * 100): "body（加权中位）",
            round(roles.get("caption", 0) * 100): "caption（最小常用）"}
    for s, n in res["sizes"]:
        A(f"| {s / 100:g} | {n} | {rmap.get(s, '')} |")
    A("")

    A("### 母版/版式固定元素（几何已换算为 DSL px）\n")
    for name, info in res["fixed"].items():
        A(f"**{name}** — 背景：{info['bg'] or '无（继承）'}")
        if info["ph"]:
            A(f"- 占位符：{', '.join(info['ph'])}")
        for p in info["pics"][:8]:
            A(f"- 图片 x={p['x']} y={p['y']} w={p['w']} h={p['h']}")
        A("")

    A("### 媒体资源清单\n")
    A("| part | 体积 | 尺寸 | 透明 | 判定 |\n|---|---|---|---|---|")
    for m in res["media"][:40]:
        guess = "背景候选" if m["px"] != "?" and max(int(v) for v in m["px"].split("x")) >= 1600 else (
            "logo/图标候选" if m["alpha"] else "素材")
        A(f"| `{os.path.basename(m['part'])}` | {m['bytes']//1024} KB | {m['px']} | "
          f"{'是' if m['alpha'] else '否'} | {guess} |")
    A("")

    A("## 🟡🔴 需要判断的项（XML 拿不到，须看缩略图或推断）\n")
    A("- 栅格与安全边距（grid.columns / gutter / margin / baseline）")
    A("- 视觉密度与留白比例（每页信息量、卡片风格）")
    A("- 版式模式：封面 / 章节扉页 / 内容页 / 结束页 各自的结构")
    A("- 图片处理方式（圆角？蒙层？裁切比例？）")
    A("- 图标风格（线性 / 面性 / 双色）与使用规则")
    A("- doAndDont（从模板一致性反推的显式约束）")
    A("")
    A("> 渲染缩略图办法（可选）：安装 LibreOffice + poppler 后运行 tencent-pptx 的 "
      "`scripts/pptx2image.py`；Windows 上也可用已安装的 PowerPoint 手动「另存为 PNG」。"
      "上述项若无法获得缩略图，则基于本报告测量值 + 模板用途推断，并在 DESIGN.md 中标注置信度。")
    p = os.path.join(outdir, "TEMPLATE-MEASURE.md")
    with open(p, "w", encoding="utf-8") as fh:
        fh.write("\n".join(L))
    return p


def write_design_md(res: dict, outdir: str) -> str:
    px_w, px_h = res["px"]
    tag = os.path.splitext(os.path.basename(res["file"]))[0]
    ea = res["ea"][0][0] if res["ea"] and res["ea"][0][0] else (res["font_scheme"].get("minorFont", {}).get("ea") or "Microsoft YaHei")
    if ea.startswith("+"):   # "+mn-ea" 之类是主题占位，换成主题实际字体
        ea = (res["font_scheme"].get("minorFont", {}).get("ea")
              or res["font_scheme"].get("minorFont", {}).get("latin") or "Microsoft YaHei")
    cs = res["color_scheme"]
    chromatic = res.get("chromatic") or res.get("useful_colors") or []

    def _lum(c: str) -> float:
        r, g, b = int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
        return 0.299 * r + 0.587 * g + 0.114 * b

    if res["stock_theme"] and chromatic:
        prim = chromatic[0]
        rest = chromatic[1:7]
        deep = min(rest, key=_lum) if rest else prim
        theme_note = "  # ⚠️ 模板用 Office 默认主题色，已改用实际高频彩色"
    else:
        prim = cs.get("accent1") or (chromatic[0] if chromatic else "#2B6CB0")
        deep = cs.get("accent2") or prim
        theme_note = "  # 🟢 主题 accent1"
    dark = (res.get("dark_ink") or [None])[0] or cs.get("dk2") or "#1A202C"
    lt = (res.get("light_surface") or [None])[0] or cs.get("lt2") or "#F5F5F8"
    roles = res["size_roles"]
    s_disp, s_body, s_cap = roles["display"], roles["body"], roles["caption"]
    from math import gcd
    iw, ih = max(1, round(px_w)), max(1, round(px_h))
    g = gcd(iw, ih) or 1
    ratio = f"{iw // g}:{ih // g}"
    if ratio == "16:9":
        ratio = "16:9"
    t = f"""---
version: "0.1"
name: "从 {tag} 抽取的设计契约"
description: "由 extract_template_design.py 自动抽取；🟢=XML 测量值，🟡=视觉估计，🔴=推断（需人工复核）"
audience: "TODO 从模板用途推断：受众 / 场景 / 气质"     # 🔴
locale: "zh-CN"                                        # 🟢

canvas:
  aspectRatio: "{ratio}"           # 🟢 注意：DSL 画布固定 1280x720，非 16:9 模板需缩放
  width: {px_w:g}                                      # 🟢 EMU/9525
  height: {px_h:g}
  scaleToDsl: {{x: {res['scale_x']}, y: {res['scale_y']}}}   # 🟢 落到 .slide 时的坐标换算系数
  safeArea:                                            # 🟡 视觉估计
    top: 24
    right: 32
    bottom: 32
    left: 32

grid:                                                  # 🔴 模板未显式声明
  columns: 12
  gutter: 16
  margin: 32
  baseline: 8

colors:
  primary: "{prim}"{theme_note}
  primary-deep: "{deep}"           # 🟢 高频深色
  on-surface: "{dark}"             # 🟢 页面高频暗色（正文/标题色）
  surface: "#FFFFFF"               # 🟢
  surface-muted: "{lt}"            # 🟢 页面高频浅色（卡片底色）
  accent: "TODO 从用色频次里挑一个与主色相距 ≥90° 的值"   # 🟡
  # 实际高频用色见 TEMPLATE-MEASURE.md；确定后收敛到 ≤4 个 hex（含中性色）

typography:
  display:
    fontFamily: '"{ea}"'
    fontWeight: 700
    fontSize: {s_disp:g}           # 🟢 模板最大字号
  title:
    fontFamily: '"{ea}"'
    fontWeight: 700
    fontSize: {max(round(s_body * 1.6), 24):g}   # 🟡 常规标题
  body:
    fontFamily: '"{ea}"'
    fontWeight: 500
    fontSize: {s_body:g}           # 🟢 模板加权中位字号
  caption:
    fontFamily: '"{ea}"'
    fontWeight: 500
    fontSize: {s_cap:g}            # 🟢 模板最小常用字号

spacing:                                               # 🟡
  xs: 4
  sm: 8
  md: 16
  lg: 24
  xl: 40
  xxl: 64

rounded:                                               # 🟡
  none: 0
  sm: 4
  md: 8
  pill: 9999

# ───── 以下字段需读模板缩略图后归纳（无缩略图则据用途推断并标 🔴）─────
imagery:            # TODO style / treatment / cornerRadius / usage / doNotUse
iconography:        # TODO style / metaphors / color / size / doNotUse
patterns:           # TODO cover / sectionDivider / contentPage / closing
motion:             # TODO declared / recommended
meta:
  totalSlides: {res['counts']['slides']}      # 🟢
  layoutCount: {res['counts']['layouts']}     # 🟢
doAndDont:          # TODO do / dont —— 从模板的一致性反推的硬约束
---
"""
    p = os.path.join(outdir, "DESIGN.md")
    with open(p, "w", encoding="utf-8") as fh:
        fh.write(t)
    return p


def find_template(target: str) -> str:
    if os.path.isfile(target):
        return target
    exts = (".pptx", ".potx", ".ppt", ".pot")
    for root, _dirs, files in os.walk(target):
        for f in sorted(files):
            if f.lower().endswith(exts) and not f.startswith("~$"):
                return os.path.join(root, f)
    raise SystemExit(f"[模板抽取] 在 {target} 下没有找到 pptx/potx 文件")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="从自有 PPT 模板抽取视觉设计契约")
    ap.add_argument("template", help="模板 .pptx/.potx 路径，或存放模板的文件夹")
    ap.add_argument("-o", "--outdir", default=None, help="输出目录（默认 <模板目录>/_提取）")
    args = ap.parse_args(argv)

    tpl = find_template(args.template)
    if tpl.lower().endswith((".ppt", ".pot")):
        raise SystemExit("[模板抽取] 旧版 .ppt 需先另存为 .pptx 再抽取")
    outdir = args.outdir or os.path.join(os.path.dirname(os.path.abspath(tpl)), "_提取")
    os.makedirs(outdir, exist_ok=True)

    res = extract(tpl, outdir)
    m = write_measure_md(res, outdir)
    d = write_design_md(res, outdir)

    print("[模板抽取] 完成")
    print(f"  模板: {tpl}")
    print(f"  画布: {res['px'][0]:g}x{res['px'][1]:g} px  比例 {res['ratio']}  "
          f"缩放->DSL x{res['scale_x']} y{res['scale_y']}")
    print(f"  页数 {res['counts']['slides']} / 版式 {res['counts']['layouts']} / 母版 {res['counts']['masters']}")
    print(f"  主色候选: " + "  ".join(f"{c}({n})" for c, n in res["colors"][:6]))
    print(f"  字体: " + "  ".join(f"{f}({n})" for f, n in res["ea"][:4]) + " | " +
          "  ".join(f"{f}({n})" for f, n in res["latin"][:4]))
    print(f"  抽取固定资产 {len(res['assets_saved'])} 个 -> {os.path.join(outdir, 'assets')}")
    print(f"  报告: {m}")
    print(f"  契约: {d}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
