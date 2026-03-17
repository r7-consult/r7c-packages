# Интерактивные уроки SmartDocs

Этот каталог хранит интерактивный курс **«Разработка макросов с нуля»**.

## Структура

- `interactive-manifest.json` — список модулей и уроков курса.
- `modules/` — hand-crafted уроки (MVP и после review).
- `cards/` — автогенерируемые карточки-черновики.
- `AUTHORING_GUIDE.md` — правила для контент-команды.

## Формат урока

Урок — это Markdown-файл с YAML frontmatter.

```yaml
---
interactive:
  version: 1
  enabled: true
  lessonId: "macro-basics-001"
  title: "Первые действия с Api.GetActiveSheet"
  targetEditor: "cell"
  task:
    goal: "..."
  starterCode: |
    (function () {
      const ws = Api.GetActiveSheet();
    })();
  beforeScript: |
    (function () {
      const ws = Api.GetActiveSheet();
      ws.GetRange("A1:C10").Clear();
    })();
  checks:
    - id: "sheet-active-exists"
      type: "sheet_exists"
      sheet: "Active"
  draft: true
---
# Теория
...
```

## Поддерживаемые проверки (MVP)

- `sheet_exists`: `{ sheet }`
- `cell_value`: `{ sheet, cell, expected, normalize? }`
- `cell_formula`: `{ sheet, cell, expected }`

## Скрипты

- Генерация 25 черновиков:

```bash
node modules/macros_ide/scripts/tools/generate-interactive-cards.js
```

- Валидация frontmatter + согласованности манифеста:

```bash
node modules/macros_ide/scripts/tools/validate-interactive-cards.js
```
