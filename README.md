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

## Архитектура Python-скрипта

4 слоя:
1. **Чтение JSONL** — парсинг cleaned.json
2. **Нормализация текста** — разделение на абзацы, сохранение переносов
3. **Выбор пресета** — preflight-оценка + postflight-валидация
4. **Генерация DOCX** — отдельная секция на каждое сообщение

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

Пути абсолютные или относительно папки с конфигом.

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
| `--no-compress` | Нет | Отключить адаптивное сжатие |

### Примеры

```bash
# Базовый запуск
python build_doc.py --input ../output --output ../output/book.docx

# Формат A5, выравнивание по центру
python build_doc.py --input ../output --output ../output/book.docx --format A5 --align center

# Без сжатия — все сообщения в одном размере
python build_doc.py --input ../output --output ../output/book.docx --no-compress

# С конфигом
python build_doc.py --input ../output --output ../output/book.docx --config ../configs/docx.json
```

### Конфиг docx.json

```json
{
  "page_format": "A4",
  "alignment": "left",
  "font_name": "Times New Roman",
  "font_size": 12,
  "no_compress": false
}
```

CLI-аргументы переопределяют значения из конфига.

### Режимы работы

**С адаптивным сжатием** (по умолчанию):

Скрипт последовательно перебирает пресеты и для каждого проверяет:
1. Preflight-оценку (суммарное количество строк с учётом абзацев и отступов)
2. Postflight-валидацию (пословный разбор с учётом длин строк и запаса 5%)

Только пройдя обе проверки, пресет считается подходящим. Если ни один не прошёл — минимальный формат + warning.

**Без сжатия** (`--no-compress`):

Все сообщения одним форматом (12pt, поля 25mm). Длинные тексты идут на несколько страниц.

## Формат cleaned.json (JSONL)

```json
{"message":"Текст первого письма"}
{"message":"Текст второго письма"}
```

Каждое сообщение — отдельная строка JSON.
