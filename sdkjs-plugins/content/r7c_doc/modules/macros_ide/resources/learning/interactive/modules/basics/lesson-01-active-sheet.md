---
interactive:
  version: 1
  enabled: true
  lessonId: "macro-basics-001"
  title: "Первые действия с Api.GetActiveSheet"
  targetEditor: "cell"
  difficulty: "easy"
  task:
    goal: "Запишите формулу в B3, чтобы умножить B1 и B2"
    steps:
      - "Получите активный лист"
      - "Установите B1=2 и B2=2"
      - "Установите формулу в B3"
    acceptance:
      - "B3 содержит формулу =B1*B2"
      - "B3 вычисляется в 4"
  starterCode: |
    (function () {
      const ws = Api.GetActiveSheet();
      // TODO: ваш код
    })();
  beforeScript: |
    (function () {
      const ws = Api.GetActiveSheet();
      ws.GetRange("A1:C10").Clear();
      ws.GetRange("B1").SetValue("2");
      ws.GetRange("B2").SetValue("2");
    })();
  checks:
    - id: "sheet-active-exists"
      type: "sheet_exists"
      sheet: "Active"
    - id: "formula-b3"
      type: "cell_formula"
      sheet: "Active"
      cell: "B3"
      expected: "=B1*B2"
    - id: "value-b3"
      type: "cell_value"
      sheet: "Active"
      cell: "B3"
      expected: 4
      normalize: "number"
  hints:
    - "Используй Api.GetActiveSheet()"
    - "Формулу можно поставить строкой, начинающейся с ="
  source:
    category: "ApiRange"
    examplePath: "modules/macros_ide/resources/learning/examples-macros/ApiRange/SetValue.md"
  draft: false
---
# Теория
`Api.GetActiveSheet()` возвращает текущий рабочий лист. Через `GetRange("A1")` получаем диапазон/ячейку и можем записывать значение через `SetValue`.

Для установки формулы используется та же операция `SetValue`, но значение должно начинаться с `=`.

Пример: `ws.GetRange("B3").SetValue("=B1*B2")`.
