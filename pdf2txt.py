#!/usr/bin/env python3
from __future__ import annotations

import argparse
import importlib.util
from pathlib import Path
from typing import Callable


def _extract_with_pymupdf(pdf_path: Path) -> str:
    import fitz

    doc = fitz.open(pdf_path)
    parts: list[str] = []
    for page in doc:
        text = page.get_text("text") or ""
        parts.append(text.rstrip() + "\n")
    doc.close()
    return "\n".join(parts).strip() + "\n"


def _extract_with_pdfminer(pdf_path: Path) -> str:
    from pdfminer.high_level import extract_text

    text = extract_text(str(pdf_path)) or ""
    return text.strip() + "\n"


def _extract_with_pypdf(pdf_path: Path) -> str:
    try:
        from pypdf import PdfReader
    except ImportError:
        from PyPDF2 import PdfReader

    reader = PdfReader(str(pdf_path))
    parts: list[str] = []
    for page in reader.pages:
        page_text = page.extract_text() or ""
        parts.append(page_text.rstrip() + "\n")
    return "\n".join(parts).strip() + "\n"


def _pick_extractor() -> tuple[str, Callable[[Path], str]]:
    def has_module(name: str) -> bool:
        try:
            return importlib.util.find_spec(name) is not None
        except ModuleNotFoundError:
            return False

    if has_module("fitz"):
        return "pymupdf", _extract_with_pymupdf

    if has_module("pdfminer.high_level"):
        return "pdfminer", _extract_with_pdfminer

    if has_module("pypdf"):
        return "pypdf", _extract_with_pypdf

    if has_module("PyPDF2"):
        return "PyPDF2", _extract_with_pypdf

    raise SystemExit(
        "缺少PDF解析依赖，请先安装其一：\n"
        "  - pip install pymupdf\n"
        "  - pip install pdfminer.six\n"
        "  - pip install pypdf\n"
        "  - pip install PyPDF2\n"
    )


def iter_pdf_files(input_dir: Path):
    for path in sorted(input_dir.glob("*.pdf")):
        if path.is_file():
            yield path


def convert_directory(input_dir: Path, output_dir: Path) -> tuple[int, int]:
    engine_name, extractor = _pick_extractor()
    output_dir.mkdir(parents=True, exist_ok=True)

    converted = 0
    skipped = 0

    for pdf_path in iter_pdf_files(input_dir):
        try:
            text = extractor(pdf_path)
        except Exception:
            skipped += 1
            continue

        out_path = output_dir / f"{pdf_path.stem}.txt"
        out_path.write_text(text, encoding="utf-8")
        converted += 1

    print(f"使用引擎: {engine_name}")
    return converted, skipped


def main() -> int:
    base_dir = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default=str(base_dir / "pdf"), help="包含 .pdf 的目录")
    parser.add_argument("--output", default=str(base_dir / "txt"), help="输出 .txt 的目录")
    args = parser.parse_args()

    input_dir = Path(args.input).expanduser().resolve()
    output_dir = Path(args.output).expanduser().resolve()

    if not input_dir.exists() or not input_dir.is_dir():
        raise SystemExit(f"输入目录不存在或不是目录: {input_dir}")

    converted, skipped = convert_directory(input_dir, output_dir)
    print(f"已转换: {converted}，跳过: {skipped}，输出目录: {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
