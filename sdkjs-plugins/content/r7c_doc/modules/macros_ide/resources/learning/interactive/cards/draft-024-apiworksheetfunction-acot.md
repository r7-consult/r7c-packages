---
interactive:
  version: 1
  enabled: true
  lessonId: "interactive-draft-024"
  title: "Черновик: ApiWorksheetFunction / ACOT"
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
            // This example shows how to return the arccotangent of a number, in radians in the range from 0 to Pi.
            
            // How to get an arccotangent of a number and display it in the worksheet.
            
            // Get a function that gets arccotangent of a number.
            
            let worksheet = Api.GetActiveSheet();
            let func = Api.GetWorksheetFunction();
            worksheet.GetRange("A1").SetValue(func.ACOT(0));
            
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
  hints:
    - "Начните с Api.GetActiveSheet()"
    - "Сверяйтесь с ожидаемым состоянием после запуска"
  source:
    category: "ApiWorksheetFunction"
    examplePath: "modules/macros_ide/resources/learning/examples-macros/ApiWorksheetFunction/ACOT.md"
  draft: true
---
# Теория
Черновая карточка создана автоматически на базе примера `ApiWorksheetFunction/ACOT`.

Перед публикацией проверьте формулировки задания, подсказки и корректность проверок.
