#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


DOCX_WORD_DIR = "word/"
DOCX_MAIN_DOCUMENT = "word/document.xml"


@dataclass(frozen=True)
class DocxText:
    paragraphs: list[str]

    def to_text(self) -> str:
        lines: list[str] = []
        for paragraph in self.paragraphs:
            lines.extend(paragraph.splitlines())
        text = "\n".join(lines)
        text = re.sub(r"[ \t]+\n", "\n", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip() + "\n"


def _read_xml_from_docx(docx_path: Path, inner_path: str) -> ET.Element | None:
    try:
        with zipfile.ZipFile(docx_path) as zf:
            try:
                data = zf.read(inner_path)
            except KeyError:
                return None
    except zipfile.BadZipFile:
        return None

    try:
        return ET.fromstring(data)
    except ET.ParseError:
        return None


def _extract_paragraph_text(p: ET.Element, ns: dict[str, str]) -> str:
    out: list[str] = []
    for node in p.iter():
        if node.tag == f"{{{ns['w']}}}t":
            if node.text:
                out.append(node.text)
        elif node.tag == f"{{{ns['w']}}}tab":
            out.append("\t")
        elif node.tag in (f"{{{ns['w']}}}br", f"{{{ns['w']}}}cr"):
            out.append("\n")
    return "".join(out).strip()


def _extract_paragraphs_from_xml(root: ET.Element) -> list[str]:
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    paragraphs: list[str] = []
    for p in root.findall(".//w:p", ns):
        text = _extract_paragraph_text(p, ns)
        paragraphs.append(text)
    return paragraphs


def extract_docx_text(docx_path: Path) -> DocxText | None:
    main = _read_xml_from_docx(docx_path, DOCX_MAIN_DOCUMENT)
    if main is None:
        return None

    paragraphs = _extract_paragraphs_from_xml(main)

    with zipfile.ZipFile(docx_path) as zf:
        extra_xmls = [
            name
            for name in zf.namelist()
            if name.startswith(DOCX_WORD_DIR)
            and (
                name.startswith("word/header")
                or name.startswith("word/footer")
                or name in ("word/footnotes.xml", "word/endnotes.xml")
            )
            and name.endswith(".xml")
        ]

    for inner in sorted(extra_xmls):
        extra = _read_xml_from_docx(docx_path, inner)
        if extra is None:
            continue
        paragraphs.extend(_extract_paragraphs_from_xml(extra))

    return DocxText(paragraphs=paragraphs)


def iter_docx_files(input_dir: Path) -> Iterable[Path]:
    for path in sorted(input_dir.glob("*.docx")):
        if path.is_file():
            yield path


def convert_directory(input_dir: Path, output_dir: Path) -> tuple[int, int]:
    output_dir.mkdir(parents=True, exist_ok=True)

    converted = 0
    skipped = 0

    for docx_path in iter_docx_files(input_dir):
        docx_text = extract_docx_text(docx_path)
        if docx_text is None:
            skipped += 1
            continue

        out_path = output_dir / f"{docx_path.stem}.txt"
        out_path.write_text(docx_text.to_text(), encoding="utf-8")
        converted += 1

    return converted, skipped


def main() -> int:
    base_dir = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--input",
        default=str(base_dir / "word"),
        help="包含 .docx 的目录",
    )
    parser.add_argument(
        "--output",
        default=str(base_dir / "txt"),
        help="输出 .txt 的目录",
    )
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
