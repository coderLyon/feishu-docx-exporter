"""PDF -> PNG（pymupdf）。逐页渲染并即时释放，避免 8GB 内存下 OOM。

用法:
    python tools/pdf2png.py <pdf路径> <输出目录> [--dpi 144] [--pages 1-12]

输出:
    <输出目录>/p001.png ...
    <输出目录>/meta.json   页数 / DPI / 每页像素尺寸 / 页面尺寸(pt)
"""
import argparse
import json
import os
import sys


def parse_pages(spec, total):
    if not spec:
        return list(range(total))
    out = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            lo = int(a or 1)
            hi = int(b or total)
        else:
            lo = hi = int(part)
        out.extend(n for n in range(lo, hi + 1) if 1 <= n <= total)
    return sorted(set(out))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("outdir")
    ap.add_argument("--dpi", type=int, default=144)
    ap.add_argument("--pages", default="")
    args = ap.parse_args()

    import pymupdf

    os.makedirs(args.outdir, exist_ok=True)
    doc = pymupdf.open(args.pdf)
    total = doc.page_count
    wanted = parse_pages(args.pages, total)

    zoom = args.dpi / 72.0
    matrix = pymupdf.Matrix(zoom, zoom)
    info = []
    for n in wanted:
        page = doc[n - 1]
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        name = "p%03d.png" % n
        path = os.path.join(args.outdir, name)
        pix.save(path)
        w, h = pix.width, pix.height
        pix = None  # 立即释放，避免累积占用
        info.append({
            "page": n,
            "file": name,
            "pxWidth": w,
            "pxHeight": h,
            "ptWidth": round(page.rect.width, 1),
            "ptHeight": round(page.rect.height, 1),
        })
        print("  %s  %dx%d" % (name, w, h), flush=True)

    meta = {
        "pdf": os.path.basename(args.pdf),
        "pageCount": total,
        "renderedPages": wanted,
        "dpi": args.dpi,
        "pages": info,
    }
    with open(os.path.join(args.outdir, "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

    doc.close()
    print("PDF2PNG_OK %d/%d pages @%dDPI" % (len(wanted), total, args.dpi))


if __name__ == "__main__":
    sys.exit(main())
