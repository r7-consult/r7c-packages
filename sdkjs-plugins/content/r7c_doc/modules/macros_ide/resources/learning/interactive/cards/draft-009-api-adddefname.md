---
interactive:
  version: 1
  enabled: true
  lessonId: "interactive-draft-009"
  title: "Черновик: Api / AddDefName"
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
            
            // Get the active worksheet
            // Получение активного листа
            let worksheet = Api.GetActiveSheet();
            
            // Set values in cells A1 and B1
            // Установка значений в ячейках A1 и B1
            worksheet.GetRange("A1").SetValue("1");
            worksheet.GetRange("B1").SetValue("2");
            
            // Add a defined name "numbers" to the range A1:B1
            // Добавление определенного имени "numbers" к диапазону A1:B1
            Api.AddDefName("numbers", "Sheet1!$A$1:$B$1");
            
            // Display a message confirming the defined name creation
            // Отображение сообщения, подтверждающего создание определенного имени
            worksheet.GetRange("A3").SetValue("We defined a name 'numbers' for a range of cells A1:B1.");
            
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
    - id: "value-a1"
      type: "cell_value"
      sheet: "Active"
      cell: "A1"
      expected: "1"
      normalize: "string"
  hints:
    - "Начните с Api.GetActiveSheet()"
    - "Сверяйтесь с ожидаемым состоянием после запуска"
  source:
    category: "Api"
    examplePath: "modules/macros_ide/resources/learning/examples-macros/Api/AddDefName.md"
  draft: true
---
# Теория
Черновая карточка создана автоматически на базе примера `Api/AddDefName`.

Перед публикацией проверьте формулировки задания, подсказки и корректность проверок.
