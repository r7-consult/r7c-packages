---
interactive:
  version: 1
  enabled: true
  lessonId: "macro-formulas-002"
  title: "Формулы: SUM по диапазону"
  targetEditor: "cell"
  difficulty: "easy"
  task:
    goal: "Заполните A1:A3 числами и поставьте в A4 формулу суммы"
    steps:
      - "Получите активный лист"
      - "Установите A1=3, A2=4, A3=5"
      - "Установите формулу =SUM(A1:A3) в A4"
    acceptance:
      - "В A4 записана формула =SUM(A1:A3)"
      - "A4 вычисляется в 12"
  starterCode: |
    (function () {
      const ws = Api.GetActiveSheet();
      // TODO: заполните A1:A3 и формулу в A4
    })();
  beforeScript: |
    (function () {
      const ws = Api.GetActiveSheet();
      ws.GetRange("A1:A10").Clear();
    })();
  checks:
    - id: "sheet-active"
      type: "sheet_exists"
      sheet: "Active"
    - id: "formula-a4"
      type: "cell_formula"
      sheet: "Active"
      cell: "A4"
      expected: "=SUM(A1:A3)"
    - id: "value-a4"
      type: "cell_value"
      sheet: "Active"
      cell: "A4"
      expected: 12
      normalize: "number"
  hints:
    - "SetValue принимает и числа, и строки"
    - "Для формулы используй =SUM(A1:A3)"
  source:
    category: "ApiRange"
    examplePath: "modules/macros_ide/resources/learning/examples-macros/ApiRange/GetFormula.md"
  draft: false
---
# Теория
В макросах ячейки заполняются по адресу: `ws.GetRange("A1")`. Формулы задаются как строка со знаком `=`.

Проверка урока сравнивает и формулу, и рассчитанное значение.
