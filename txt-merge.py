from __future__ import annotations

from pathlib import Path


def _read_text_with_fallback(path: Path) -> str:
    for encoding in ("utf-8", "utf-8-sig", "gb18030"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def merge_txt_files(src_dir: Path, out_file: Path) -> int:
    if not src_dir.exists() or not src_dir.is_dir():
        raise FileNotFoundError(f"源目录不存在或不是目录: {src_dir}")

    out_file.parent.mkdir(parents=True, exist_ok=True)

    txt_files = sorted(
        (p for p in src_dir.iterdir() if p.is_file() and p.suffix.lower() == ".txt"),
        key=lambda p: p.name,
    )

    merged_parts: list[str] = []
    for i, p in enumerate(txt_files):
        content = _read_text_with_fallback(p)
        if i != 0 and merged_parts and not merged_parts[-1].endswith("\n"):
            merged_parts.append("\n")
        if i != 0:
            merged_parts.append("\n")
        merged_parts.append(content)

    out_file.write_text("".join(merged_parts), encoding="utf-8")
    return len(txt_files)


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    src_dir = base_dir / "txt"
    out_file = base_dir / "merge" / "merged.txt"
    count = merge_txt_files(src_dir=src_dir, out_file=out_file)
    print(f"已合并 {count} 个文件 -> {out_file}")


if __name__ == "__main__":
    main()
