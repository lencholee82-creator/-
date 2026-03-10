from __future__ import annotations

import argparse
from pathlib import Path

from .scanner import export_to_excel, scan_pdf_files


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="扫描目录中的 PDF 文件并导出为 Excel")
    parser.add_argument("input_dir", type=Path, help="要扫描的目录")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("pdf_filenames.xlsx"),
        help="输出 Excel 文件路径（默认: ./pdf_filenames.xlsx）",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="递归扫描子目录中的 PDF 文件",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    records = scan_pdf_files(args.input_dir, recursive=args.recursive)
    export_to_excel(records, args.output)

    print(f"已扫描到 {len(records)} 个 PDF 文件。")
    print(f"Excel 已生成: {args.output.resolve()}")


if __name__ == "__main__":
    main()
