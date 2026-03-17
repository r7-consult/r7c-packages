---
interactive:
  version: 1
  enabled: true
  lessonId: "interactive-draft-017"
  title: "Черновик: Api / attachEvent"
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
            // Инициализация API R7 Office
            const api = Api;
            if (!api) {
                throw new Error('R7 Office API not available');
            }
            
            // Get the active worksheet and a range
            // Получение активного листа и диапазона
            let worksheet = Api.GetActiveSheet();
            let range = worksheet.GetRange("A1");
            
            // Set a value in cell A1 to trigger the onWorksheetChange event
            // Установка значения в ячейке A1 для вызова события onWorksheetChange
            range.SetValue("1");
            
            // Attach an event listener for onWorksheetChange
            // Прикрепление слушателя событий для onWorksheetChange
            Api.attachEvent("onWorksheetChange", function(range){
                console.log("onWorksheetChange");
                console.log(range.GetAddress());
            });
            
            // Success notification
            // Уведомление об успешном выполнении
            console.log('Macro executed successfully');
            
        } catch (error) {
            // Error handling
            // Обработка ошибок
            console.error('Macro execution failed:', error.message);
            // Optional: Show error to user in cell A1 if API is available
            // Опционально: Показать ошибку пользователю в ячейке A1, если API доступен
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
    category: "Api"
    examplePath: "modules/macros_ide/resources/learning/examples-macros/Api/attachEvent.md"
  draft: true
---
# Теория
Черновая карточка создана автоматически на базе примера `Api/attachEvent`.

Перед публикацией проверьте формулировки задания, подсказки и корректность проверок.
