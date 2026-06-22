# MailParser

Парсинг EML-писем → JSON → DOCX.

## Структура проекта

```
MailParser/
├── configs/
│   ├── parser.json      — настройки Go-парсера (входная/выходная папки)
│   └── docx.json        — настройки DOCX (шрифт, формат, выравнивание)
├── input/eml/           — сюда кладёшь .eml файлы
├── output/              — результат: cleaned.json, book.docx
├── go-parser/           — Go: EML → JSON
└── py-docx/             — Python: JSON → DOCX
```

## Запуск Go-парсера

```bash
cd go-parser
go run cmd/mailparser/main.go
```

Парсер читает `configs/parser.json` (дефолтный путь). Чтобы указать другой конфиг:

```bash
go run cmd/mailparser/main.go -config /путь/к/parser.json
```

Результат: `output/cleaned.json` — JSONL файл (один JSON-объект на строку).

### Конфиг parser.json

```json
{
  "input_dir": "input/eml",
  "output_dir": "output"
}
```

Для корректной работы необходимо указать абсолютные пути или относительные пути от папки с конфигом.

## Запуск Python-скрипта

```bash
cd py-docx
pip install -r requirements.txt
python build_doc.py --input ../output --output ../output/book.docx
```

### Аргументы

| Флаг | Обязательный | Описание |
|------|-------------|----------|
| `--input` | Да | Папка с cleaned.json |
| `--output` | Да | Путь к выходному .docx |
| `--format` | Нет | A4 или A5 (по умолчанию A4) |
| `--align` | Нет | left, center, right, justify (по умолчанию left) |
| `--config` | Нет | Путь к docx.json |

### Примеры

```bash
# Базовый запуск
python build_doc.py --input ../output --output ../output/book.docx

# Формат A5, выравнивание по центру
python build_doc.py --input ../output --output ../output/book.docx --format A5 --align center

# С конфигом
python build_doc.py --input ../output --output ../output/book.docx --config ../configs/docx.json
```

### Конфиг docx.json

```json
{
  "page_format": "A4",
  "alignment": "left",
  "font_name": "Times New Roman",
  "font_size": 12
}
```

CLI-аргументы переопределяют значения из конфига.

### Адаптивное сжатие

Если сообщение не помещается на одну страницу, скрипт автоматически ужимает оформление:

1. Уменьшает шрифт (12 → 11 → 10 → 9 → 8pt)
2. Сужает поля (25 → 20 → 18 → 15 → 12mm)
3. Уменьшает межстрочный интервал (1.15 → 1.1 → 1.0)
4. Сокращает отступы между абзацами

Если после всех пресетов текст всё равно не влезает — скрипт предупредит в консоли и использует минимальный формат.

## Формат cleaned.json (JSONL)

```json
{"message":"Текст первого письма"}
{"message":"Текст второго письма"}
```

Каждое сообщение — отдельная строка JSON.
