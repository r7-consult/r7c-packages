---
interactive:
  version: 1
  enabled: true
  lessonId: "interactive-draft-023"
  title: "Черновик: ApiWorksheet / AddWordArt"
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
            // This example adds a Text Art object to the sheet with the parameters specified.
            
            // How to add a word art to the worksheet specifying its properties, color, size, etc.
            
            // Insert a word art to the worksheet.
            
            let worksheet = Api.GetActiveSheet();
            let textProps = Api.CreateTextPr();
            textProps.SetFontSize(72);
            textProps.SetBold(true);
            textProps.SetCaps(true);
            textProps.SetColor(51, 51, 51, false);
            textProps.SetFontFamily("Comic Sans MS");
            let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
            let stroke = Api.CreateStroke(1 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51)));
            worksheet.AddWordArt(textProps, "R7 Office", "textArchUp", fill, stroke, 0, 100 * 36000, 20 * 36000, 0, 2, 2 * 36000, 3 * 36000);
            
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
    examplePath: "modules/macros_ide/resources/learning/examples-macros/ApiWorksheet/AddWordArt.md"
  draft: true
---
# Теория
Черновая карточка создана автоматически на базе примера `ApiWorksheet/AddWordArt`.

Перед публикацией проверьте формулировки задания, подсказки и корректность проверок.
