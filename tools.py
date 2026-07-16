"""Модульная конвертация DOCX в HTML-фрагмент без CSS и внешней обёртки.

Главные точки входа:

* :func:`convert_docx_to_html` — один DOCX в один HTML-файл;
* :func:`convert_directory` — все DOCX из указанной папки;
* :func:`document_to_html`, :func:`paragraph_to_html` и :func:`table_to_html`
  — отдельные функции для последующего расширения обработки, например фото.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Iterator
from dataclasses import dataclass
from html import escape
from pathlib import Path
import re
from typing import TypeVar

import docx
from docx.document import Document as WordDocument
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import Length
from docx.styles.style import BaseStyle
from docx.table import _Cell, Table
from docx.text.hyperlink import Hyperlink
from docx.text.paragraph import Paragraph
from docx.text.run import Run


__all__ = [
    "ConversionOptions",
    "ImageExporter",
    "convert_directory",
    "convert_docx_to_html",
    "document_to_html",
    "paragraph_to_html",
    "table_to_html",
    "write_html",
]


@dataclass(frozen=True)
class ConversionOptions:
    """Настройки преобразования HTML-фрагмента.

    Если пользовательский стиль называется «Заголовок» или ``Heading`` без
    номера, он будет записан указанным уровнем: по умолчанию ``<h3>``.
    Стандартные стили Word Heading 1…Heading 6 сохраняют свой уровень.
    """

    default_heading_level: int = 3

    def __post_init__(self) -> None:
        if not 1 <= self.default_heading_level <= 6:
            raise ValueError("default_heading_level должен быть от 1 до 6")


@dataclass(frozen=True)
class _ListInfo:
    tag: str
    level: int


_StyleValue = TypeVar("_StyleValue")
_HEADING_RE = re.compile(r"(?:heading|заголовок)\s*([1-9])?", re.IGNORECASE)
_IMAGE_CAPTION_RE = re.compile(r"^\s*(?:рисунок|рис\.)\s*", re.IGNORECASE)
_DEFAULT_IMAGE_BASE_URL = "https://lms.cbt-center.ru/wp-content/uploads/2026/07"
_INVALID_FILE_NAME_CHARACTERS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_VML_IMAGE_DATA_TAG = "{urn:schemas-microsoft-com:vml}imagedata"
_IMAGE_CONTENT_TYPE_EXTENSIONS = {
    "image/bmp": ".bmp",
    "image/gif": ".gif",
    "image/jpeg": ".jpeg",
    "image/png": ".png",
    "image/tiff": ".tiff",
    "image/x-emf": ".emf",
    "image/x-wmf": ".wmf",
}
_HIGHLIGHT_COLORS = {
    WD_COLOR_INDEX.BLACK: "#000000",
    WD_COLOR_INDEX.BLUE: "#0000ff",
    WD_COLOR_INDEX.TURQUOISE: "#00ffff",
    WD_COLOR_INDEX.BRIGHT_GREEN: "#00ff00",
    WD_COLOR_INDEX.PINK: "#ff00ff",
    WD_COLOR_INDEX.RED: "#ff0000",
    WD_COLOR_INDEX.YELLOW: "#ffff00",
    WD_COLOR_INDEX.WHITE: "#ffffff",
    WD_COLOR_INDEX.DARK_BLUE: "#000080",
    WD_COLOR_INDEX.TEAL: "#008080",
    WD_COLOR_INDEX.GREEN: "#008000",
    WD_COLOR_INDEX.VIOLET: "#800080",
    WD_COLOR_INDEX.DARK_RED: "#800000",
    WD_COLOR_INDEX.DARK_YELLOW: "#808000",
    WD_COLOR_INDEX.GRAY_50: "#808080",
    WD_COLOR_INDEX.GRAY_25: "#c0c0c0",
}


@dataclass
class ImageExporter:
    """Сохраняет изображения одного DOCX и формирует теги для HTML.

    Экземпляр предназначен для одной конвертации, поэтому ``counter`` всегда
    начинается с нуля для каждого DOCX-файла.
    """

    document_name: str
    output_directory: Path
    base_url: str = _DEFAULT_IMAGE_BASE_URL
    counter: int = 0

    def __post_init__(self) -> None:
        self.output_directory.mkdir(parents=True, exist_ok=True)

    def export_from_run(self, run: Run) -> list[str]:
        """Сохранить изображения из одного Run и вернуть теги ``img``."""

        image_tags: list[str] = []
        for relationship_id in _image_relationship_ids(run):
            image_part = run.part.related_parts.get(relationship_id)
            if image_part is None:
                continue

            self.counter += 1
            extension = _image_extension(str(image_part.partname), image_part.content_type)
            file_name = f"рис_{self.counter}_{self.document_name}{extension}"
            image_path = self.output_directory / file_name
            image_path.write_bytes(image_part.blob)

            image_url = f"{self.base_url.rstrip('/')}/{file_name}"
            image_tags.append(
                f'<center><img src="{image_url}" '
                f'alt="рисунок {self.counter}"></center>'
            )
        return image_tags


def convert_docx_to_html(
    docx_path: str | Path,
    html_path: str | Path | None = None,
    *,
    options: ConversionOptions | None = None,
) -> Path:
    """Конвертировать один ``.docx`` и вернуть путь к созданному HTML-файлу.

    Результат — HTML-фрагмент, содержащий только теги содержимого. Если путь
    назначения не указан, файл создаётся рядом с DOCX с расширением ``.html``.
    """

    source = Path(docx_path)
    if source.suffix.lower() != ".docx":
        raise ValueError(f"Ожидался файл .docx, получен: {source}")
    if not source.is_file():
        raise FileNotFoundError(f"DOCX-файл не найден: {source}")

    destination = Path(html_path) if html_path is not None else source.with_suffix(".html")
    document = docx.Document(source)
    folder_document_name = _safe_file_name(source.stem)
    image_exporter = ImageExporter(
        document_name=_image_file_name_part(source.stem),
        output_directory=destination.parent / f"Рисунки_{folder_document_name}",
    )
    html_fragment = document_to_html(
        document,
        options=options,
        image_exporter=image_exporter,
    )
    write_html(html_fragment, destination)
    return destination


def convert_directory(
    input_dir: str | Path,
    output_dir: str | Path | None = None,
    *,
    recursive: bool = False,
    options: ConversionOptions | None = None,
) -> list[Path]:
    """Конвертировать каждый DOCX в папке и вернуть пути созданных HTML.

    Если ``output_dir`` не задан, HTML записывается рядом с исходным DOCX.
    Временные файлы Word, начинающиеся с ``~$``, игнорируются.
    """

    source_dir = Path(input_dir)
    if not source_dir.is_dir():
        raise NotADirectoryError(f"Папка с DOCX не найдена: {source_dir}")

    destination_dir = Path(output_dir) if output_dir is not None else None
    if destination_dir is not None:
        destination_dir.mkdir(parents=True, exist_ok=True)

    docx_files: Iterable[Path]
    docx_files = source_dir.rglob("*.docx") if recursive else source_dir.glob("*.docx")
    results: list[Path] = []
    for source in sorted(docx_files):
        if source.name.startswith("~$"):
            continue

        if destination_dir is None:
            destination = source.with_suffix(".html")
        elif recursive:
            destination = destination_dir / source.relative_to(source_dir).with_suffix(".html")
        else:
            destination = destination_dir / source.with_suffix(".html").name
        results.append(convert_docx_to_html(source, destination, options=options))
    return results


def write_html(html_fragment: str, path: str | Path) -> Path:
    """Записать HTML-фрагмент в UTF-8 и вернуть путь к файлу."""

    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(html_fragment, encoding="utf-8")
    return destination


def document_to_html(
    document: WordDocument,
    *,
    options: ConversionOptions | None = None,
    image_exporter: ImageExporter | None = None,
) -> str:
    """Преобразовать DOCX в HTML-фрагмент без ``html``/``body``.

    Передайте ``image_exporter``, чтобы сохранять встроенные изображения и
    добавлять к фрагменту HTML-теги ``img``.
    """

    return _blocks_to_html(
        _iter_block_items(document),
        options or ConversionOptions(),
        image_exporter,
    )


def paragraph_to_html(
    paragraph: Paragraph,
    *,
    options: ConversionOptions | None = None,
    image_exporter: ImageExporter | None = None,
) -> str:
    """Преобразовать один абзац в ``p`` либо ``h1``…``h6`` с его текстом."""

    actual_options = options or ConversionOptions()
    heading_level = _heading_level(paragraph.style.name, actual_options.default_heading_level)
    tag = f"h{heading_level}" if heading_level is not None else "p"
    text, image_tags = _inline_content_to_html(paragraph, image_exporter)
    html_parts: list[str] = []
    force_center = heading_level in {2, 3} or _is_image_caption(paragraph)
    if text:
        paragraph_html = f"<{tag}{_alignment_attribute(paragraph)}>{text}</{tag}>"
        html_parts.append(_center_html_if_needed(paragraph_html, paragraph, force_center))
    elif not image_tags:
        paragraph_html = f"<{tag}{_alignment_attribute(paragraph)}><br></{tag}>"
        html_parts.append(_center_html_if_needed(paragraph_html, paragraph, force_center))
    html_parts.extend(image_tags)
    return "\n".join(html_parts)


def table_to_html(
    table: Table,
    *,
    options: ConversionOptions | None = None,
    image_exporter: ImageExporter | None = None,
) -> str:
    """Преобразовать таблицу DOCX в теги ``table``, ``tr`` и ``td``."""

    actual_options = options or ConversionOptions()
    rows: list[str] = []
    for row in table.rows:
        cells: list[str] = []
        for cell in row.cells:
            text = _blocks_to_html(
                _iter_block_items(cell),
                actual_options,
                image_exporter,
            )
            cells.append(f"<td{_cell_attributes(cell)}>{text}</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return '<table border="1" cellpadding="6" cellspacing="0">' + "".join(rows) + "</table>"


def _iter_block_items(parent: WordDocument | _Cell) -> Iterator[Paragraph | Table]:
    """Выдать абзацы и таблицы в том порядке, в каком они стоят в DOCX."""

    parent_element = parent.element.body if isinstance(parent, WordDocument) else parent._tc
    for child in parent_element.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def _blocks_to_html(
    blocks: Iterable[Paragraph | Table],
    options: ConversionOptions,
    image_exporter: ImageExporter | None,
) -> str:
    output: list[str] = []
    list_stack: list[_ListInfo] = []

    def close_lists() -> None:
        while list_stack:
            output.append(f"</li></{list_stack.pop().tag}>")

    for block in blocks:
        if isinstance(block, Table):
            close_lists()
            output.append(
                table_to_html(
                    block,
                    options=options,
                    image_exporter=image_exporter,
                )
            )
            continue

        list_info = _get_list_info(block)
        if list_info is None:
            close_lists()
            output.append(
                paragraph_to_html(
                    block,
                    options=options,
                    image_exporter=image_exporter,
                )
            )
            continue

        level = min(list_info.level, len(list_stack))
        list_info = _ListInfo(list_info.tag, level)
        while len(list_stack) - 1 > level:
            output.append(f"</li></{list_stack.pop().tag}>")

        if len(list_stack) > level and list_stack[level].tag != list_info.tag:
            output.append(f"</li></{list_stack.pop().tag}>")

        if len(list_stack) == level:
            output.append(f"<{list_info.tag}>")
            list_stack.append(list_info)
            output.append("<li>")
        else:
            output.append("</li><li>")
        text, image_tags = _inline_content_to_html(block, image_exporter)
        list_item_html = text or "<br>"
        output.append(_center_html_if_needed(list_item_html, block, _is_image_caption(block)))
        output.extend(image_tags)

    close_lists()
    return "\n".join(output)


def _heading_level(style_name: str, default_level: int) -> int | None:
    match = _HEADING_RE.search(style_name)
    if match is None:
        return None
    return min(int(match.group(1)), 6) if match.group(1) is not None else default_level


def _get_list_info(paragraph: Paragraph) -> _ListInfo | None:
    numbering = _numbering_values(paragraph)
    if numbering is not None:
        number_id, level = numbering
        return _ListInfo(_numbering_tag(paragraph, number_id, level), level)

    style_name = paragraph.style.name.casefold()
    if "bullet" in style_name or "маркирован" in style_name:
        return _ListInfo("ul", 0)
    if "list" in style_name or "список" in style_name or "нумер" in style_name:
        return _ListInfo("ol", 0)
    return None


def _numbering_values(paragraph: Paragraph) -> tuple[str, int] | None:
    """Вернуть ``(numId, ilvl)`` из абзаца или наследуемого стиля списка."""

    styles = [paragraph.style]
    while styles[-1].base_style is not None:
        styles.append(styles[-1].base_style)

    for properties in [paragraph._p.pPr, *(style.element.pPr for style in styles)]:
        if properties is None:
            continue
        num_pr = properties.find(qn("w:numPr"))
        if num_pr is None:
            continue
        num_id = num_pr.find(qn("w:numId"))
        if num_id is None:
            continue
        number_id = num_id.get(qn("w:val"))
        if not number_id or number_id == "0":
            continue
        ilvl = num_pr.find(qn("w:ilvl"))
        level = int(ilvl.get(qn("w:val"), "0")) if ilvl is not None else 0
        return number_id, level
    return None


def _numbering_tag(paragraph: Paragraph, number_id: str, level: int) -> str:
    """Вернуть ``ul`` для маркеров, во всех прочих случаях — ``ol``."""

    numbering = paragraph.part.numbering_part.element
    abstract_id: str | None = None
    for num in numbering.findall(qn("w:num")):
        if num.get(qn("w:numId")) == number_id:
            abstract_num_id = num.find(qn("w:abstractNumId"))
            if abstract_num_id is not None:
                abstract_id = abstract_num_id.get(qn("w:val"))
            break
    if abstract_id is None:
        return "ol"

    for abstract_num in numbering.findall(qn("w:abstractNum")):
        if abstract_num.get(qn("w:abstractNumId")) != abstract_id:
            continue
        for current_level in abstract_num.findall(qn("w:lvl")):
            if current_level.get(qn("w:ilvl")) != str(level):
                continue
            num_format = current_level.find(qn("w:numFmt"))
            return "ul" if num_format is not None and num_format.get(qn("w:val")) == "bullet" else "ol"
    return "ol"


def _image_relationship_ids(run: Run) -> list[str]:
    """Получить rId встроенных изображений в порядке их появления в Run."""

    relationship_ids: list[str] = []
    for blip in run._element.findall(".//" + qn("a:blip")):
        relationship_id = blip.get(qn("r:embed"))
        if relationship_id and relationship_id not in relationship_ids:
            relationship_ids.append(relationship_id)
    for image_data in run._element.findall(".//" + _VML_IMAGE_DATA_TAG):
        relationship_id = image_data.get(qn("r:id"))
        if relationship_id and relationship_id not in relationship_ids:
            relationship_ids.append(relationship_id)
    return relationship_ids


def _image_extension(part_name: str, content_type: str) -> str:
    extension = Path(part_name).suffix
    if extension:
        return extension
    return _IMAGE_CONTENT_TYPE_EXTENSIONS.get(content_type, ".bin")


def _safe_file_name(name: str) -> str:
    cleaned_name = _INVALID_FILE_NAME_CHARACTERS.sub("_", name).rstrip(". ")
    return cleaned_name or "document"


def _image_file_name_part(document_name: str) -> str:
    """Подготовить имя документа для файла изображения: пробелы → дефисы."""

    safe_name = _safe_file_name(document_name)
    return re.sub(r"\s+", "-", safe_name)


def _inline_content_to_html(
    paragraph: Paragraph,
    image_exporter: ImageExporter | None,
) -> tuple[str, list[str]]:
    """Получить текст и HTML-теги изображений из одного абзаца Word."""

    output: list[str] = []
    image_tags: list[str] = []
    for item in paragraph.iter_inner_content():
        if isinstance(item, Run):
            output.append(_run_to_html(item, paragraph))
            if image_exporter is not None:
                image_tags.extend(image_exporter.export_from_run(item))
            continue

        hyperlink = item
        link_parts: list[str] = []
        for run in hyperlink.runs:
            link_parts.append(_run_to_html(run, paragraph))
            if image_exporter is not None:
                image_tags.extend(image_exporter.export_from_run(run))
        link_text = "".join(link_parts)
        if hyperlink.url:
            url = escape(hyperlink.url, quote=True)
            output.append(f'<a href="{url}">{link_text}</a>')
        else:
            output.append(link_text)
    return "".join(output), image_tags


def _run_to_html(run: Run, paragraph: Paragraph) -> str:
    """Обернуть текст одного Word Run в соответствующие HTML-теги."""

    text = escape(run.text).replace("\t", "&emsp;").replace("\n", "<br>")
    if not text:
        return ""

    font = run.font
    font_name = _resolve_font_value(font.name, run.style, paragraph.style, lambda style: style.font.name)
    font_size = _resolve_font_value(font.size, run.style, paragraph.style, lambda style: style.font.size)
    font_color = _resolve_font_value(
        font.color.rgb,
        run.style,
        paragraph.style,
        lambda style: style.font.color.rgb,
    )
    highlight = _resolve_font_value(
        font.highlight_color,
        run.style,
        paragraph.style,
        lambda style: style.font.highlight_color,
    )

    font_attributes: list[str] = []
    if font_name is not None:
        font_attributes.append(f'face="{escape(font_name, quote=True)}"')
    if font_size is not None:
        font_attributes.append(f'size="{_html_font_size(font_size)}"')
    if font_color is not None:
        font_attributes.append(f'color="#{font_color}"')
    if font_attributes:
        text = f"<font {' '.join(font_attributes)}>{text}</font>"

    highlight_color = _HIGHLIGHT_COLORS.get(highlight)
    if highlight_color is not None:
        text = f"<mark>{text}</mark>"
    if _resolve_font_value(font.bold, run.style, paragraph.style, lambda style: style.font.bold):
        text = f"<strong>{text}</strong>"
    if _resolve_font_value(font.italic, run.style, paragraph.style, lambda style: style.font.italic):
        text = f"<em>{text}</em>"
    if _resolve_font_value(font.underline, run.style, paragraph.style, lambda style: style.font.underline):
        text = f"<u>{text}</u>"
    if _resolve_font_value(font.strike, run.style, paragraph.style, lambda style: style.font.strike) or _resolve_font_value(
        font.double_strike,
        run.style,
        paragraph.style,
        lambda style: style.font.double_strike,
    ):
        text = f"<s>{text}</s>"
    if _resolve_font_value(font.superscript, run.style, paragraph.style, lambda style: style.font.superscript):
        text = f"<sup>{text}</sup>"
    elif _resolve_font_value(font.subscript, run.style, paragraph.style, lambda style: style.font.subscript):
        text = f"<sub>{text}</sub>"
    return text


def _resolve_font_value(
    direct_value: _StyleValue | None,
    run_style: BaseStyle | None,
    paragraph_style: BaseStyle,
    style_value: Callable[[BaseStyle], _StyleValue | None],
) -> _StyleValue | None:
    """Получить прямое значение шрифта или значение из цепочки Word-стилей."""

    if direct_value is not None:
        return direct_value
    for first_style in (run_style, paragraph_style):
        current_style = first_style
        while current_style is not None:
            value = style_value(current_style)
            if value is not None:
                return value
            current_style = current_style.base_style
    return None


def _alignment_attribute(paragraph: Paragraph) -> str:
    alignment = _resolve_paragraph_alignment(paragraph)
    alignment_map = {
        WD_ALIGN_PARAGRAPH.LEFT: "left",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
        WD_ALIGN_PARAGRAPH.DISTRIBUTE: "justify",
    }
    html_alignment = alignment_map.get(alignment)
    return f' align="{html_alignment}"' if html_alignment is not None else ""


def _center_html_if_needed(html: str, paragraph: Paragraph, force: bool = False) -> str:
    """Обернуть центрированный или принудительно центрированный текст."""

    if force or _resolve_paragraph_alignment(paragraph) == WD_ALIGN_PARAGRAPH.CENTER:
        return f"<center>{html}</center>"
    return html


def _is_image_caption(paragraph: Paragraph) -> bool:
    """Распознать подпись к рисунку по её тексту: «Рисунок …» или «Рис. …»."""

    return _IMAGE_CAPTION_RE.match(paragraph.text) is not None


def _resolve_paragraph_alignment(paragraph: Paragraph) -> WD_ALIGN_PARAGRAPH | None:
    direct_alignment = paragraph.paragraph_format.alignment
    if direct_alignment is not None:
        return direct_alignment
    style = paragraph.style
    while style is not None:
        alignment = style.paragraph_format.alignment
        if alignment is not None:
            return alignment
        style = style.base_style
    return None


def _cell_attributes(cell: _Cell) -> str:
    fill = _shading_fill(cell._tc.tcPr)
    return f' bgcolor="{fill}"' if fill is not None else ""


def _shading_fill(properties: object | None) -> str | None:
    if properties is None:
        return None
    shading = properties.find(qn("w:shd"))
    if shading is None:
        return None
    fill = shading.get(qn("w:fill"))
    if fill and fill.lower() not in {"auto", "none"}:
        return f"#{fill}"
    return None


def _html_font_size(value: Length) -> int:
    """Приблизительно перевести Word points в допустимый ``font size`` 1…7."""

    points = value.pt
    if points <= 8:
        return 1
    if points <= 10:
        return 2
    if points <= 12:
        return 3
    if points <= 14:
        return 4
    if points <= 18:
        return 5
    if points <= 24:
        return 6
    return 7
