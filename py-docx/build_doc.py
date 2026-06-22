"""
Генерация DOCX из cleaned.json.

Каждое сообщение — отдельная страница. Если текст не помещается,
скрипт автоматически ужимает оформление (шрифт, поля, интервалы).
"""

import argparse
import json
import math
import os
import sys

from docx import Document
from docx.shared import Mm, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ── Константы ────────────────────────────────────────────────

# Размеры страниц в миллиметрах
PAGE_SIZES = {
    "A4": (210, 297),
    "A5": (148, 210),
}

# Соответствие строковых имён выравнивания объектам python-docx
ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

# Пресеты сжатия: от стандартного (читаемого) к минимальному.
# Перебираются по порядку, пока текст не влезет на одну страницу.
COMPRESS_PRESETS = [
    {"font_size": 12, "margins_mm": 25, "line_spacing": 1.15, "before": 0, "after": 6},
    {"font_size": 11, "margins_mm": 20, "line_spacing": 1.15, "before": 0, "after": 4},
    {"font_size": 10, "margins_mm": 18, "line_spacing": 1.15, "before": 0, "after": 3},
    {"font_size": 10, "margins_mm": 15, "line_spacing": 1.1,  "before": 0, "after": 2},
    {"font_size": 9,  "margins_mm": 15, "line_spacing": 1.0,  "before": 0, "after": 2},
    {"font_size": 8,  "margins_mm": 12, "line_spacing": 1.0,  "before": 0, "after": 1},
]

# ── Загрузка конфига и аргументов ────────────────────────────

def load_config(config_path):
    """Читает docx.json, возвращает dict (или пустой dict если файла нет)."""
    if not os.path.exists(config_path):
        return {}
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def parse_args():
    """Парсит CLI-аргументы. CLI переопределяет значения из конфига."""
    parser = argparse.ArgumentParser(
        description="Генерация DOCX из cleaned.json",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры:
  python build_doc.py --input ../output --output ../output/book.docx
  python build_doc.py --input ../output --output ../output/book.docx --format A5 --align left
        """,
    )
    parser.add_argument("--input", required=True, help="Папка с cleaned.json")
    parser.add_argument("--output", required=True, help="Путь к выходному .docx")
    parser.add_argument("--format", choices=["A4", "A5"], default=None, help="Формат страницы (по умолчанию A4)")
    parser.add_argument("--align", choices=["left", "center", "right", "justify"], default=None,
                        help="Выравнивание (по умолчанию left)")
    parser.add_argument("--config", default=None, help="Путь к docx.json")
    return parser.parse_args()


def resolve_settings(args):
    """Объединяет CLI-аргументы и конфиг. CLI имеет приоритет."""
    cfg = {}
    if args.config:
        cfg = load_config(args.config)

    return {
        "page_format": args.format or cfg.get("page_format", "A4"),
        "alignment": args.align or cfg.get("alignment", "left"),
        "font_name": cfg.get("font_name", "Times New Roman"),
        "font_size": cfg.get("font_size", 12),
    }

# ── Чтение JSONL ─────────────────────────────────────────────

def read_messages(json_path):
    """
    Читает cleaned.json в формате JSONL (один JSON-объект на строку).
    Возвращает список текстовых сообщений.
    """
    if not os.path.exists(json_path):
        print(f"ОШИБКА: Файл не найден: {json_path}")
        sys.exit(1)

    messages = []
    with open(json_path, "r", encoding="utf-8") as f:
        for line_num, line in enumerate(f, 1):
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
                text = obj.get("message", "")
                if text:
                    messages.append(text)
            except json.JSONDecodeError as e:
                print(f"Предупреждение: строка {line_num} — невалидный JSON: {e}")
    return messages

# ── Оценка заполнения страницы ───────────────────────────────

def estimate_chars_per_line(font_size_pt, page_width_mm, margins_mm):
    """
    Оценивает, сколько символов помещается в одну строку.
    Использует среднюю ширину символа для пропорционального шрифта (~0.5 * размер).
    1 pt = 0.3528 mm.
    """
    usable_mm = page_width_mm - 2 * margins_mm
    avg_char_width_mm = font_size_pt * 0.5 * 0.3528
    if avg_char_width_mm <= 0:
        avg_char_width_mm = 1.0
    return max(1, int(usable_mm / avg_char_width_mm))


def estimate_lines_for_text(text, font_size_pt, page_width_mm, margins_mm):
    """Сколько строк нужно для размещения текста."""
    chars_per_line = estimate_chars_per_line(font_size_pt, page_width_mm, margins_mm)
    total_chars = len(text)
    lines = math.ceil(total_chars / chars_per_line)
    return max(1, lines)


def max_lines_on_page(font_size_pt, line_spacing, page_height_mm, margins_mm):
    """Сколько строк максимального размера помещается на страницу."""
    usable_height_mm = page_height_mm - 2 * margins_mm
    line_height_mm = font_size_pt * 0.3528 * line_spacing
    if line_height_mm <= 0:
        line_height_mm = 1.0
    return max(1, int(usable_height_mm / line_height_mm))

# ── Выбор пресета ────────────────────────────────────────────

def find_best_preset(text, page_w, page_h, alignment):
    """
    Перебирает пресеты от самого читаемого к самому компактному.
    Возвращает первый, при котором текст помещается на страницу.
    Если не один не влезает — возвращает минимальный.
    """
    for preset in COMPRESS_PRESETS:
        lines_needed = estimate_lines_for_text(
            text, preset["font_size"], page_w, preset["margins_mm"]
        )
        lines_available = max_lines_on_page(
            preset["font_size"], preset["line_spacing"], page_h, preset["margins_mm"]
        )
        if lines_needed <= lines_available:
            return preset

    # Текст не влез ни в один пресет — берём минимальный
    return COMPRESS_PRESETS[-1]

# ── Форматирование ───────────────────────────────────────────

def format_paragraph(paragraph, preset, font_name, alignment):
    """Применяет к абзацу настройки из пресета: интервалы, выравнивание, шрифт."""
    pf = paragraph.paragraph_format
    pf.space_before = Pt(preset["before"])
    pf.space_after = Pt(preset["after"])
    pf.line_spacing = preset["line_spacing"]

    if alignment in ALIGNMENTS:
        paragraph.alignment = ALIGNMENTS[alignment]

    for run in paragraph.runs:
        run.font.size = Pt(preset["font_size"])
        run.font.name = font_name


def format_section(section, margins_mm):
    """Устанавливает поля страницы (все четыре стороны одинаковые)."""
    section.top_margin = Mm(margins_mm)
    section.bottom_margin = Mm(margins_mm)
    section.left_margin = Mm(margins_mm)
    section.right_margin = Mm(margins_mm)

# ── Генерация документа ──────────────────────────────────────

def create_document(messages, settings):
    """
    Создаёт DOCX: каждое сообщение — отдельная страница.
    Возвращает (документ, кол-во уместившихся, кол-во сжатых).
    """
    doc = Document()
    page_w, page_h = PAGE_SIZES[settings["page_format"]]
    alignment = settings["alignment"]
    font_name = settings["font_name"]

    section = doc.sections[0]
    section.page_width = Mm(page_w)
    section.page_height = page_h

    fitted_count = 0
    overflow_count = 0

    for i, text in enumerate(messages):
        # Подбираем пресет, при котором текст влезет
        preset = find_best_preset(text, page_w, page_h, alignment)
        format_section(section, preset["margins_mm"])

        # Создаём абзац с текстом
        paragraph = doc.add_paragraph()
        run = paragraph.add_run(text)
        run.font.name = font_name
        run.font.size = Pt(preset["font_size"])

        format_paragraph(paragraph, preset, font_name, alignment)

        # Проверяем, влез ли текст фактически
        lines_needed = estimate_lines_for_text(
            text, preset["font_size"], page_w, preset["margins_mm"]
        )
        lines_available = max_lines_on_page(
            preset["font_size"], preset["line_spacing"], page_h, preset["margins_mm"]
        )

        if lines_needed <= lines_available:
            fitted_count += 1
        else:
            overflow_count += 1
            print(f"  Сообщение {i+1}: не поместилось ({lines_needed} строк > {lines_available} макс.), использован минимальный формат")

        # Разрыв страницы после каждого сообщения (кроме последнего)
        if i < len(messages) - 1:
            doc.add_page_break()

    return doc, fitted_count, overflow_count

# ── Точка входа ──────────────────────────────────────────────

def main():
    args = parse_args()
    settings = resolve_settings(args)

    # Формируем путь к cleaned.json во входной папке
    json_path = os.path.join(args.input, "cleaned.json")
    print(f"Чтение: {json_path}")
    messages = read_messages(json_path)
    print(f"Найдено сообщений: {len(messages)}")

    if not messages:
        print("Нет сообщений для обработки")
        sys.exit(0)

    print(f"Формат: {settings['page_format']}, выравнивание: {settings['alignment']}")
    print(f"Шрифт: {settings['font_name']}, начальный размер: {settings['font_size']}pt")
    print()

    doc, fitted, overflow = create_document(messages, settings)

    # Создаём папку для выходного файла, если нужно
    output_dir = os.path.dirname(args.output)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    doc.save(args.output)

    # Итоговая статистика
    print()
    print(f"Обработано: {len(messages)} сообщений")
    if fitted > 0:
        print(f"  Уместилось: {fitted}")
    if overflow > 0:
        print(f"  Сжато до минимума: {overflow}")
    print(f"Результат: {args.output}")


if __name__ == "__main__":
    main()
