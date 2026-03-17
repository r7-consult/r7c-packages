---
interactive:
  version: 1
  enabled: true
  lessonId: "interactive-draft-019"
  title: "Черновик: ApiWorksheet / AddShape"
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
            // This example adds a shape to the sheet with the parameters specified.
            
            // How to add a shape to the worksheet.
            
            // Insert a flowchart shape to the worksheet.
            
            let worksheet = Api.GetActiveSheet();
            let gradientStop1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);
            let gradientStop2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);
            let fill = Api.CreateLinearGradientFill([gradientStop1, gradientStop2], 5400000);
            let stroke = Api.CreateStroke(0, Api.CreateNoFill());
            worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
            
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
    category: "ApiWorksheet"
    examplePath: "modules/macros_ide/resources/learning/examples-macros/ApiWorksheet/AddShape.md"
  draft: true
---
# Теория
Черновая карточка создана автоматически на базе примера `ApiWorksheet/AddShape`.

Перед публикацией проверьте формулировки задания, подсказки и корректность проверок.
