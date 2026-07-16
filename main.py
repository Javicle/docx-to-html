"""Консольная точка входа для конвертации DOCX в HTML-фрагмент."""

from __future__ import annotations

import argparse
from pathlib import Path

from tools import ConversionOptions, convert_directory, convert_docx_to_html


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Конвертировать DOCX в HTML-фрагмент без CSS и HTML-обёртки."
    )
    parser.add_argument(
        "source",
        type=Path,
        help="Исходный .docx или папка с .docx при --directory.",
    )
    parser.add_argument(
        "destination",
        type=Path,
        help="Путь к HTML-файлу или папка для HTML при --directory.",
    )
    parser.add_argument(
        "-d",
        "--directory",
        action="store_true",
        help="Обработать все DOCX из source; destination считается папкой вывода.",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="Искать DOCX также во вложенных папках (только вместе с --directory).",
    )
    parser.add_argument(
        "--default-heading-level",
        type=int,
        default=3,
        choices=range(1, 7),
        metavar="1..6",
        help="Уровень для пользовательского стиля «Заголовок» без номера (по умолчанию: 3).",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    if args.recursive and not args.directory:
        parser.error("--recursive можно использовать только вместе с --directory")

    options = ConversionOptions(default_heading_level=args.default_heading_level)
    try:
        if not args.directory:
            html_path = convert_docx_to_html(
                args.source, args.destination, options=options
            )
            print(f"Создан: {html_path}")
            return 0

        generated_files = convert_directory(
            args.source,
            args.destination,
            recursive=args.recursive,
            options=options,
        )
    except (FileNotFoundError, NotADirectoryError, ValueError) as error:
        parser.error(str(error))

    if not generated_files:
        print("В указанной папке DOCX-файлы не найдены.")
        return 0

    for html_path in generated_files:
        print(f"Создан: {html_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
