---
interactive:
  version: 1
  enabled: true
  lessonId: "interactive-draft-003"
  title: "Черновик: ApiWorksheet / AddChart"
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
            // This example creates a chart of the specified type from the selected data range of the sheet.
            
            // How to add chart to the worksheet.
            
            // Create a chart using data from a range from a worksheet.
            
            let worksheet = Api.GetActiveSheet();
            worksheet.GetRange("B1").SetValue(2014);
            worksheet.GetRange("C1").SetValue(2015);
            worksheet.GetRange("D1").SetValue(2016);
            worksheet.GetRange("A2").SetValue("Projected Revenue");
            worksheet.GetRange("A3").SetValue("Estimated Costs");
            worksheet.GetRange("B2").SetValue(200);
            worksheet.GetRange("B3").SetValue(250);
            worksheet.GetRange("C2").SetValue(240);
            worksheet.GetRange("C3").SetValue(260);
            worksheet.GetRange("D2").SetValue(280);
            worksheet.GetRange("D3").SetValue(280);
            let chart = worksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);
            chart.SetTitle("Financial Overview", 13);
            let fill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
            chart.SetSeriesFill(fill, 0, false);
            fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
            chart.SetSeriesFill(fill, 1, false);
            
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
    - id: "value-b1"
      type: "cell_value"
      sheet: "Active"
      cell: "B1"
      expected: 2014
      normalize: "number"
  hints:
    - "Начните с Api.GetActiveSheet()"
    - "Сверяйтесь с ожидаемым состоянием после запуска"
  source:
    category: "ApiWorksheet"
    examplePath: "modules/macros_ide/resources/learning/examples-macros/ApiWorksheet/AddChart.md"
  draft: true
---
# Теория
Черновая карточка создана автоматически на базе примера `ApiWorksheet/AddChart`.

Перед публикацией проверьте формулировки задания, подсказки и корректность проверок.
