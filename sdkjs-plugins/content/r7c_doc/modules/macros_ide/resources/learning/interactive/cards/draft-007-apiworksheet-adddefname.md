---
interactive:
  version: 1
  enabled: true
  lessonId: "interactive-draft-007"
  title: "Черновик: ApiWorksheet / AddDefName"
  targetEditor: "cell"
  difficulty: "easy"
  task:
    goal: "Запустите макрос и получите ожидаемый результат в таблице"
    steps:
      - "Изучите пример и его цель"
      - "Доработайте код в редакторе"
      - "Запустите макрос и пройдите проверки"
    acceptance:
      - "Макрос выполняется без ошибок"
      - "Проверки урока пройдены"
  starterCode: |
    (function() {
        'use strict';
        
        try {
            // Initialize R7 Office API
            const api = Api;
            if (!api) {
                throw new Error('R7 Office API not available');
            }
            
            // Original code enhanced with error handling:
            // This example adds a new name to the worksheet.
            
            // How to change a name of the worksheet range.
            
            // Name a range from a worksheet.
            
            let worksheet = Api.GetActiveSheet();
            worksheet.GetRange("A1").SetValue("1");
            worksheet.GetRange("B1").SetValue("2");
            worksheet.AddDefName("numbers", "Sheet1!$A$1:$B$1");
            worksheet.GetRange("A3").SetValue("We defined a name 'numbers' for a range of cells A1:B1.");
            
            // Success notification
            console.log('Macro executed successfully');
            
        } catch (error) {
            console.error('Macro execution failed:', error.message);
            // Optional: Show error to user
            if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
                const sheet = Api.GetActiveSheet();
                if (sheet) {
                    sheet.GetRange('A1').SetValue('Error: ' + error.message);
                }
            }
        }
    })();
    
  beforeScript: |
    (function () {
      const ws = Api.GetActiveSheet();
      ws.GetRange("A1:Z30").Clear();
    })();
  checks:
    - id: "sheet-active-exists"
      type: "sheet_exists"
      sheet: "Active"
    - id: "value-a1"
      type: "cell_value"
      sheet: "Active"
      cell: "A1"
      expected: "1"
      normalize: "string"
  hints:
    - "Начните с Api.GetActiveSheet()"
    - "Сверяйтесь с ожидаемым состоянием после запуска"
  source:
    category: "ApiWorksheet"
    examplePath: "modules/macros_ide/resources/learning/examples-macros/ApiWorksheet/AddDefName.md"
  draft: true
---
# Теория
Черновая карточка создана автоматически на базе примера `ApiWorksheet/AddDefName`.

Перед публикацией проверьте формулировки задания, подсказки и корректность проверок.
