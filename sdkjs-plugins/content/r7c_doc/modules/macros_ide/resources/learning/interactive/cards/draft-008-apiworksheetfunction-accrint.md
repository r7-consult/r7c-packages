---
interactive:
  version: 1
  enabled: true
  lessonId: "interactive-draft-008"
  title: "Черновик: ApiWorksheetFunction / ACCRINT"
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
            // This example shows how to return the accrued interest for a security that pays periodic interest.
            
            // How to get an accrued interest for a security that pays periodic interest.
            
            // Get a function that gets accrued interest for a security.
            
            let worksheet = Api.GetActiveSheet();
            let func = Api.GetWorksheetFunction();
            worksheet.GetRange("A1").SetValue(func.ACCRINT("1/1/2018", "6/25/2018", "10/15/2018", "3.50%", 1000, 2));
            
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
    examplePath: "modules/macros_ide/resources/learning/examples-macros/ApiWorksheetFunction/ACCRINT.md"
  draft: true
---
# Теория
Черновая карточка создана автоматически на базе примера `ApiWorksheetFunction/ACCRINT`.

Перед публикацией проверьте формулировки задания, подсказки и корректность проверок.
