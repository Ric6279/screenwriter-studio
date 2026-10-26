#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""PDF 取字：有文本层直接读，扫描页用本机 RapidOCR 识别。

这是从 course-kit 的 pdf_ocr.py 里裁出来的“知识库上传”版本，只保留编剧工坊
需要的能力：把一份 PDF（剧本/编剧书）取成纯文本，供 ingest_knowledge.py
生成摘录。没有章节解析、没有逐页大缓存——上传一份 30 MB 的扫描剧本时，
先抽前 N 页（默认 120）够用，整本 OCR 动辄十几分钟，不适合卡在上传请求里。

依赖复用 course-kit/.pylibs（pymupdf + rapidocr_onnxruntime），不另装一套。
可用环境变量覆盖：
  SCREENWRITER_OCR_LIBS   依赖目录（默认 <工作区>/course-kit/.pylibs）
  SCREENWRITER_OCR_PAGES  最多处理多少页（默认 120，0 = 全部）
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from pathlib import Path

MAX_CHARS = 300_000
DEFAULT_PAGES = 120
TEXT_PAGE_MIN_CHARS = 40
NUM_ONLY_RE = re.compile(r"^[\s\d\-—–.·]+$")


def _ensure_libs() -> Path | None:
    """把 course-kit 的 .pylibs 挂到 sys.path；找不到就返回 None。"""
    candidates = [
        os.environ.get("SCREENWRITER_OCR_LIBS"),
        str(Path(__file__).resolve().parents[4] / "course-kit" / ".pylibs"),
        str(Path(__file__).resolve().parents[3] / "course-kit" / ".pylibs"),
    ]
    for cand in candidates:
        if cand and Path(cand).is_dir():
            p = str(Path(cand))
            if p not in sys.path:
                sys.path.insert(0, p)
            return Path(p)
    return None


def _load_engine(intra: int = 4):
    try:
        from rapidocr_onnxruntime import RapidOCR
    except ImportError:  # 新版包名
        from rapidocr import RapidOCR

    return RapidOCR(intra_op_num_threads=intra)


def _ocr_image(engine, image: bytes) -> list[str]:
    out = engine(image)
    if isinstance(out, tuple):  # 老版返回 (result, elapse)
        result = out[0]
    elif hasattr(out, "boxes"):  # 新版对象
        result = [[box, text, score] for box, text, score in zip(out.boxes, out.txts, out.scores)]
    else:
        result = out
    lines: list[tuple[float, float, str]] = []
    for item in result or []:
        try:
            box, text, score = item[0], item[1], item[2]
        except (TypeError, IndexError):
            continue
        text = str(text).strip()
        if not text or NUM_ONLY_RE.match(text):
            continue
        try:
            polys = [(float(p[0]), float(p[1])) for p in box]
        except (TypeError, ValueError):
            continue
        if not polys:
            continue
        x0 = min(p[0] for p in polys)
        y0 = min(p[1] for p in polys)
        lines.append((y0, x0, text))
    lines.sort(key=lambda t: (round(t[0] / 8.0), t[1]))
    return [t[2] for t in lines]


def _join_lines(parts: list[str]) -> str:
    out = ""
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if not out:
            out = part
        elif re.search(r"[A-Za-z0-9,;:)\]]$", out) and re.match(r"^[A-Za-z0-9(\[]", part):
            out += " " + part
        elif out.endswith("-") and part[:1].islower():
            out = out[:-1] + part
        else:
            out += part
    return out


def extract_text(pdf: str, max_pages: int = DEFAULT_PAGES) -> dict:
    lib_dir = _ensure_libs()
    try:
        import pymupdf
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError(
            "缺少 pymupdf。把 course-kit/.pylibs 放到默认位置，或设置 SCREENWRITER_OCR_LIBS。"
        ) from exc

    doc = pymupdf.open(pdf)
    total = doc.page_count
    # max_pages 限制的是“要 OCR 的页数”，不是总页数：文字版 PDF 仍然整本读，
    # 扫描件只认前 N 页，避免一次上传卡十几分钟。
    ocr_limit = total if max_pages <= 0 else max_pages
    chunks: list[str] = []
    chars = 0
    text_pages = 0
    ocr_pages = 0
    ocr_skipped = 0
    engine = None
    started = time.time()

    try:
        for index in range(total):
            page = doc[index]
            text = (page.get_text("text") or "").strip()
            if len(text) >= TEXT_PAGE_MIN_CHARS:
                text_pages += 1
            elif ocr_pages < ocr_limit:
                if engine is None:
                    engine = _load_engine()
                pix = page.get_pixmap(dpi=200, colorspace=pymupdf.csRGB, alpha=False)
                lines = _ocr_image(engine, pix.tobytes("png"))
                text = _join_lines(lines)
                ocr_pages += 1
            else:
                ocr_skipped += 1
                continue
            if not text:
                continue
            chunks.append(text)
            chars += len(text)
            if chars >= MAX_CHARS:
                break
    finally:
        doc.close()

    out = "\n\n".join(chunks)[:MAX_CHARS]
    return {
        "ok": bool(out.strip()),
        "text": out,
        "totalPages": total,
        "textPages": text_pages,
        "ocrPages": ocr_pages,
        "ocrSkipped": ocr_skipped,
        "ocrLimit": ocr_limit,
        "chars": len(out),
        "libDir": str(lib_dir) if lib_dir else None,
        "seconds": round(time.time() - started, 1),
    }


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--outDir", default=".")
    ap.add_argument("--pages", type=int, default=0, help="最多处理多少页，默认按环境变量/120")
    args = ap.parse_args(argv)

    max_pages = args.pages or int(os.environ.get("SCREENWRITER_OCR_PAGES") or DEFAULT_PAGES)
    src = Path(args.pdf)
    try:
        result = extract_text(str(src), max_pages)
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"ok": False, "message": str(exc)}, ensure_ascii=True))
        return 1

    if not result["ok"]:
        print(json.dumps({
            **{k: v for k, v in result.items() if k != "text"},
            "message": "一页也没取到文字（扫描太糊或整本都是插图）。可调高 DPI 或换文件。",
        }, ensure_ascii=True))
        return 2

    out_dir = Path(args.outDir)
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"{src.stem}.ocr.txt"
    header = (
        f"# OCR 取字 · {src.name}\n\n"
        f"> 共 {result['totalPages']} 页；文字层 {result['textPages']} 页，"
        f"OCR {result['ocrPages']} 页"
        + (f"（其余 {result['ocrSkipped']} 页扫描页超出单次上限，未识别）" if result["ocrSkipped"] else "")
        + f"；{result['chars']} 字；{result['seconds']} 秒。\n\n"
    )
    out.write_text(header + result["text"] + "\n", encoding="utf-8")
    print(json.dumps({
        **{k: v for k, v in result.items() if k != "text"},
        "out": str(out),
        "method": "pymupdf+rapidocr" if result["ocrPages"] else "pymupdf",
    }, ensure_ascii=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
