---
interactive:
  version: 1
  enabled: true
  lessonId: "interactive-draft-004"
  title: "Черновик: ApiWorksheetFunction / ABS"
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
            // This example shows how to get absolute value of a number.
            
            // How to add absolute value to the worksheet.
            
            // Get a function that gets absolute value.
            
            let worksheet = Api.GetActiveSheet();
            let func = Api.GetWorksheetFunction();
            worksheet.GetRange("A1").SetValue(func.ABS(-123.14));
            
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
    examplePath: "modules/macros_ide/resources/learning/examples-macros/ApiWorksheetFunction/ABS.md"
  draft: true
---
# Теория
Черновая карточка создана автоматически на базе примера `ApiWorksheetFunction/ABS`.

Перед публикацией проверьте формулировки задания, подсказки и корректность проверок.
