"""PDF 级实测断言（比 XML 更强：验证的是 Word 真实渲染结果）。

用法:
    python tools/pdf-probe.py <pdf> [在线锚点meta.json]

产出:
    <pdf同目录>/probe.json   字号/颜色/行距实测 + 关键页定位
    <pdf同目录>/anchors.json 在线标题锚点 → Word 页码（供 make-manifest 配对）

判据:
    H1/H2/H3 实际字号 = 16.0 / 14.0 / 12.0 pt（±0.2）
    标题行内字号集合 size==1（编号与文字同号）
    标题行颜色 = 0x1F2329 或 000000（黑色）
    正文行距 Δ/size ∈ [1.60, 1.80]（雅黑单倍 ≈1.714em）
"""
import json
import os
import re
import sys

EXPECT_HEADING_PT = {1: 16.0, 2: 14.0, 3: 12.0}
BLACK = {0x1F2329, 0x000000, 0x0}
BLUE = 0x3370FF


def lines_of(page):
    out = []
    for blk in page.get_text("dict")["blocks"]:
        if blk.get("type") != 0:
            continue
        for ln in blk.get("lines", []):
            spans = ln.get("spans", [])
            if not spans:
                continue
            text = "".join(s["text"] for s in spans).strip()
            if not text:
                continue
            out.append({
                "text": text,
                "sizes": sorted({round(s["size"], 1) for s in spans}),
                "colors": sorted({int(s["color"]) & 0xFFFFFF for s in spans}),
                "fonts": sorted({s["font"] for s in spans}),
                "y0": round(ln["bbox"][1], 1),
                "x0": round(ln["bbox"][0], 1),
            })
    return out


def main():
    import pymupdf

    pdf = sys.argv[1]
    anchor_meta = sys.argv[2] if len(sys.argv) > 2 else ""
    outdir = os.path.dirname(os.path.abspath(pdf))
    doc = pymupdf.open(pdf)

    all_lines = []
    for pno in range(doc.page_count):
        for ln in lines_of(doc[pno]):
            ln["page"] = pno + 1
            all_lines.append(ln)

    # 标题判定：文本形如 "1. 文档说明" / "1.1 编写说明" / "6.3.1 判定规则"
    heads = []
    for ln in all_lines:
        t = ln["text"]
        m = re.match(r"^(\d+(?:\.\d+)*)[.．]\s*(\S.*)$", t)
        if not m or len(t) > 60:
            continue
        level = m.group(1).count(".") + 1
        if level > 3:
            continue
        heads.append({**ln, "level": level, "num": m.group(1), "title": m.group(2)})

    checks = []

    # 标题字号：不能用「正则匹配所有编号行」——列表项（12pt）与表格文字（10.6pt）
    # 会混进样本把结论带偏。改为按在线锚点（真实标题文本）逐条定位后核对。
    def line_of(page_no, needle):
        target = re.sub(r"\s+", "", needle)
        for ln in all_lines:
            if ln["page"] == page_no and target and target in re.sub(r"\s+", "", ln["text"]):
                return ln
        return None

    anchor_lines = []
    if anchor_meta and os.path.exists(anchor_meta):
        for a in json.load(open(anchor_meta, encoding="utf-8")).get("anchors", []):
            text = a.get("text", "").replace("\u200b", "").strip()
            # H2/H3 的编号打在文本里（"1.1 编写说明"）；
            # H1 的编号由飞书渲染在文本之外，锚点文本就是纯标题（"文档说明"）→ 视为一级
            # 注意用 [.．]?（可选）：写成 [.．] 会让「1.1 编写说明」贪心回溯成
            # 前缀「1」+ 正文「1 编写说明」，层级被误判为一级
            m = re.match(r"^(\d+(?:\.\d+)*)[.．]?\s*(\S.*)$", text)
            lvl = (m.group(1).count(".") + 1) if m else 1
            if lvl > 3:
                continue
            for pno in range(doc.page_count):
                ln = line_of(pno + 1, text)
                if ln:
                    anchor_lines.append({**ln, "level": lvl})
                    break

    if anchor_lines:
        for lvl, want in EXPECT_HEADING_PT.items():
            hs = [h for h in anchor_lines if h["level"] == lvl and h["sizes"]]
            if not hs:
                checks.append({"id": f"C{lvl}", "desc": f"H{lvl} 实测字号", "ok": False, "detail": "未找到"})
                continue
            same = sum(1 for h in hs if len(h["sizes"]) == 1)
            black = sum(1 for h in hs if all(c in BLACK for c in h["colors"]))
            top = max({round(s, 1): 0 for h in hs for s in h["sizes"]}.keys())
            sizes = {round(s, 1): sum(1 for h in hs for x in h["sizes"] if round(x, 1) == round(s, 1))
                     for h in hs for s in h["sizes"]}
            checks.append({
                "id": f"C{lvl}",
                "desc": f"H{lvl} 实测字号 = {want}pt",
                "ok": abs(top - want) <= 0.2,
                "detail": f"实测 {top}pt（分布 {sizes}），共 {len(hs)} 处",
            })
            checks.append({
                "id": f"C{lvl}b",
                "desc": f"H{lvl} 编号与文字同号",
                "ok": same == len(hs),
                "detail": f"{same}/{len(hs)} 处行内字号唯一",
            })
            checks.append({
                "id": f"C{lvl}c",
                "desc": f"H{lvl} 为黑色",
                "ok": black == len(hs),
                "detail": f"{black}/{len(hs)} 处黑色",
            })
    else:
        checks.append({"id": "C0", "desc": "标题字号核对", "ok": False,
                       "detail": "未提供在线锚点 meta.json，跳过（传第二个参数启用）"})

    # 蓝色残留：只允许出现在超链接
    blue_pages = {}
    for ln in all_lines:
        if BLUE in ln["colors"]:
            blue_pages.setdefault(ln["page"], []).append(ln["text"][:24])
    checks.append({
        "id": "C4",
        "desc": "蓝色仅出现在超链接（≤3 处）",
        "ok": len(blue_pages) <= 3,
        "detail": f"蓝字页: { {k: v[:2] for k, v in blue_pages.items()} }",
    })

    # 正文行距：找连续同字号同 x0 的行
    ratios = []
    for pno in range(doc.page_count):
        ls = [l for l in all_lines if l["page"] == pno + 1 and l["sizes"] and abs(l["sizes"][0] - 12.0) <= 0.3]
        ls.sort(key=lambda l: (l["x0"], l["y0"]))
        for a, b in zip(ls, ls[1:]):
            if abs(a["x0"] - b["x0"]) > 2:
                continue
            d = b["y0"] - a["y0"]
            if 8 < d < 40:
                ratios.append(d / a["sizes"][0])
    ratio = round(sum(ratios) / len(ratios), 3) if ratios else 0
    checks.append({
        "id": "C5",
        "desc": "正文行距 Δ/size ∈ [1.60,1.80]（单倍）",
        "ok": 1.60 <= ratio <= 1.80,
        "detail": f"实测 {ratio}（样本 {len(ratios)}）",
    })

    # 关键页定位
    # 注意：PDF 文本层会插入额外空格（如「表中400、401、500 指响应体…」），
    # 直接做子串匹配会漏；统一去掉空白后再比，避免定位到错误页
    def page_of(needle):
        target = re.sub(r"\s+", "", needle)
        for ln in all_lines:
            if target and target in re.sub(r"\s+", "", ln["text"]):
                return ln["page"]
        return None

    anchors = []
    if anchor_meta and os.path.exists(anchor_meta):
        meta = json.load(open(anchor_meta, encoding="utf-8"))
        for a in meta.get("anchors", []):
            t = a.get("text", "").replace("\u200b", "").strip()
            if not t:
                continue
            hit = None
            for pno in range(doc.page_count):
                if doc[pno].search_for(t):
                    hit = pno + 1
                    break
            if hit:
                anchors.append({"text": t, "page": hit, "onlineY": a.get("y", 0)})

    out = {
        "pdf": os.path.basename(pdf),
        "pageCount": doc.page_count,
        "checks": checks,
        "keyPages": {
            "2.4 接口流程图": page_of("2.4 接口流程图"),
            "6.2 试验状态流转": page_of("6.2 试验状态流转"),
            "5.1 异常场景": page_of("5.1 异常场景"),
            "引用块首条": page_of("表中400、401、500指响应体中的业务code"),
            "附件A": page_of("附件A"),
            "附件C": page_of("附件C"),
            "建表语句caption": page_of("PLAN_WO_JDZJ 原建表语句"),
        },
        "anchors": anchors,
    }
    json.dump(out, open(os.path.join(outdir, "probe.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    if anchors:
        json.dump(anchors, open(os.path.join(outdir, "anchors.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=2)

    print(f"页数: {doc.page_count}")
    for c in checks:
        print(f"{'✓' if c['ok'] else '✗'} {c['id']} {c['desc']}  → {c['detail']}")
    print("关键页:", json.dumps(out["keyPages"], ensure_ascii=False))
    print("锚点配对:", len(anchors))
    doc.close()


if __name__ == "__main__":
    main()
