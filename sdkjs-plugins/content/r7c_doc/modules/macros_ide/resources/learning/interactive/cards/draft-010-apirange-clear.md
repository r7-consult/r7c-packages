---
interactive:
  version: 1
  enabled: true
  lessonId: "interactive-draft-010"
  title: "Черновик: ApiRange / Clear"
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
            // This example clears the range.
            
            // How to clear a content of a range.
            
            // Get a range and remove everything from it.
            
            let worksheet = Api.GetActiveSheet();
            let range = worksheet.GetRange("A1:B1");
            range.SetValue("1");
            range.Clear();
            worksheet.GetRange("A2").SetValue("The range A1:B1 was just cleared.");
            
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
    - id: "value-a2"
      type: "cell_value"
      sheet: "Active"
      cell: "A2"
      expected: "The range A1:B1 was just cleared."
      normalize: "string"
  hints:
    - "Начните с Api.GetActiveSheet()"
    - "Сверяйтесь с ожидаемым состоянием после запуска"
  source:
    category: "ApiRange"
    examplePath: "modules/macros_ide/resources/learning/examples-macros/ApiRange/Clear.md"
  draft: true
---
# Теория
Черновая карточка создана автоматически на базе примера `ApiRange/Clear`.

Перед публикацией проверьте формулировки задания, подсказки и корректность проверок.
