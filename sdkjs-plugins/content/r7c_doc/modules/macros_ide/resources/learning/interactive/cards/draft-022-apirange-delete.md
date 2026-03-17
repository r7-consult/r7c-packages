---
interactive:
  version: 1
  enabled: true
  lessonId: "interactive-draft-022"
  title: "Черновик: ApiRange / Delete"
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
            // This example deletes the Range object.
            
            // How to remove a range from the worksheet.
            
            // Get a range from the worksheet and delete it specifying the direction.
            
            let worksheet = Api.GetActiveSheet();
            worksheet.GetRange("B4").SetValue("1");
            worksheet.GetRange("C4").SetValue("2");
            worksheet.GetRange("D4").SetValue("3");
            worksheet.GetRange("C5").SetValue("5");
            let range = worksheet.GetRange("C4");
            range.Delete("up");
            
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
    - id: "value-b4"
      type: "cell_value"
      sheet: "Active"
      cell: "B4"
      expected: "1"
      normalize: "string"
  hints:
    - "Начните с Api.GetActiveSheet()"
    - "Сверяйтесь с ожидаемым состоянием после запуска"
  source:
    category: "ApiRange"
    examplePath: "modules/macros_ide/resources/learning/examples-macros/ApiRange/Delete.md"
  draft: true
---
# Теория
Черновая карточка создана автоматически на базе примера `ApiRange/Delete`.

Перед публикацией проверьте формулировки задания, подсказки и корректность проверок.
