---
interactive:
  version: 1
  enabled: true
  lessonId: "interactive-draft-002"
  title: "Черновик: ApiRange / AddComment"
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
            // This example adds a comment to the range.
            
            // How to comment a range.
            
            // Get a range from the worksheet, add a comment to it and then show the comments text.
            
            let worksheet = Api.GetActiveSheet();
            let range = worksheet.GetRange("A1");
            range.SetValue("1");
            range.AddComment("This is just a number.");
            worksheet.GetRange("A3").SetValue("The comment was added to the cell A1.");
            worksheet.GetRange("A4").SetValue("Comment: " + range.GetComment().GetText());
            
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
    - id: "value-a3"
      type: "cell_value"
      sheet: "Active"
      cell: "A3"
      expected: "The comment was added to the cell A1."
      normalize: "string"
  hints:
    - "Начните с Api.GetActiveSheet()"
    - "Сверяйтесь с ожидаемым состоянием после запуска"
  source:
    category: "ApiRange"
    examplePath: "modules/macros_ide/resources/learning/examples-macros/ApiRange/AddComment.md"
  draft: true
---
# Теория
Черновая карточка создана автоматически на базе примера `ApiRange/AddComment`.

Перед публикацией проверьте формулировки задания, подсказки и корректность проверок.
