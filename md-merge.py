from __future__ import annotations

import argparse
import gzip
import re
from pathlib import Path


def _read_text_with_fallback(path: Path) -> str:
    for encoding in ("utf-8", "utf-8-sig", "gb18030"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


_BLANK_LINES_RE = re.compile(r"\n{3,}")


def _minify_markdown(text: str) -> str:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.rstrip() for line in normalized.split("\n")]
    normalized = "\n".join(lines)
    normalized = _BLANK_LINES_RE.sub("\n\n", normalized)
    if not normalized.endswith("\n"):
        normalized += "\n"
    return normalized


def merge_md_files(src_dir: Path, out_file: Path, *, minify: bool = False, gzip_output: bool = False) -> int:
    if not src_dir.exists() or not src_dir.is_dir():
        raise FileNotFoundError(f"源目录不存在或不是目录: {src_dir}")

    out_file.parent.mkdir(parents=True, exist_ok=True)

    md_files = sorted(
        (p for p in src_dir.rglob("*.md") if p.is_file()),
        key=lambda p: p.relative_to(src_dir).as_posix(),
    )

    merged_parts: list[str] = []
    for i, p in enumerate(md_files):
        rel = p.relative_to(src_dir).as_posix()
        content = _read_text_with_fallback(p)
        if i != 0:
            merged_parts.append("\n\n" if minify else "\n\n---\n\n")
        merged_parts.append(f"## {rel}\n\n")
        merged_parts.append(content)
        if merged_parts and not merged_parts[-1].endswith("\n"):
            merged_parts.append("\n")

    merged_text = "".join(merged_parts)
    if minify:
        merged_text = _minify_markdown(merged_text)

    if gzip_output:
        if out_file.suffix.lower() != ".gz":
            out_file = out_file.with_name(out_file.name + ".gz")
        with gzip.open(out_file, "wb", compresslevel=9) as f:
            f.write(merged_text.encode("utf-8"))
    else:
        out_file.write_text(merged_text, encoding="utf-8")
    return len(md_files)


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, default=base_dir / "md")
    parser.add_argument("--output", type=Path, default=base_dir / "merge" / "merged.md")
    parser.add_argument("--minify", action="store_true")
    parser.add_argument("--gzip", action="store_true")
    args = parser.parse_args()

    from datetime import datetime

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = args.output.stem
    suffix = args.output.suffix
    args.output = args.output.with_name(f"{stem}_{timestamp}{suffix}")


    count = merge_md_files(
        src_dir=args.input,
        out_file=args.output,
        minify=args.minify,
        gzip_output=args.gzip,
    )

    out_path = args.output
    if args.gzip and out_path.suffix.lower() != ".gz":
        out_path = out_path.with_name(out_path.name + ".gz")
    print(f"已合并 {count} 个文件 -> {out_path}")


if __name__ == "__main__":
    main()
