---
interactive:
  version: 1
  enabled: true
  lessonId: "macro-worksheets-003"
  title: "Работа с листами: создание и активация"
  targetEditor: "cell"
  difficulty: "medium"
  task:
    goal: "Создайте лист 'Отчет' и запишите в A1 текст 'Готово'"
    steps:
      - "Создайте лист с именем Отчет"
      - "Сделайте его активным"
      - "Запишите в A1 значение Готово"
    acceptance:
      - "Лист Отчет существует"
      - "В A1 листа Отчет записано Готово"
  starterCode: |
    (function () {
      // TODO: создайте и заполните лист "Отчет"
    })();
  beforeScript: |
    (function () {
      // В этом уроке состояние не очищаем, чтобы не затронуть пользовательские листы
    })();
  checks:
    - id: "sheet-report-exists"
      type: "sheet_exists"
      sheet: "Отчет"
    - id: "sheet-report-value"
      type: "cell_value"
      sheet: "Отчет"
      cell: "A1"
      expected: "Готово"
      normalize: "string"
  hints:
    - "Для создания листа обычно используют Api.AddSheet(\"Отчет\")"
    - "Лист можно получить через Api.GetSheet(\"Отчет\")"
    - "Значение записывается через GetRange(...).SetValue(...)"
  source:
    category: "ApiWorksheet"
    examplePath: "modules/macros_ide/resources/learning/examples-macros/ApiWorksheet/SetActive.md"
  draft: false
---
# Теория
Для работы с несколькими листами используйте API книги:

- создание листа;
- поиск листа по имени;
- установка активного листа;
- запись значений через `GetRange`.

Проверка в этом уроке смотрит только факт существования листа и значение в одной ячейке.
