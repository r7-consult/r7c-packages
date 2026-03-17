---
interactive:
  version: 1
  enabled: true
  lessonId: "interactive-draft-021"
  title: "Черновик: Api / ClearCustomFunctions"
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
            
            // Add a custom function library and a custom function named ADD
            // Добавление библиотеки пользовательских функций и пользовательской функции с именем ADD
            Api.AddCustomFunctionLibrary("LibraryName", function(){
                /**
                 * Function that returns the argument
                 * Функция, которая возвращает аргумент
                 * @customfunction
                 * @param {any} first First argument. (Первый аргумент.)
                 * @returns {any} second Second argument. (Второй аргумент.)
                 */
                Api.AddCustomFunction(function ADD(first, second) {
                    return first + second;
                });
            });
            
            // Get the active worksheet
            // Получение активного листа
            let worksheet = Api.GetActiveSheet();
            
            // Use the custom function in cell A1
            // Использование пользовательской функции в ячейке A1
            worksheet.GetRange("A1").SetValue("=ADD(1, 2)");
            
            // Clear all custom functions
            // Удаление всех пользовательских функций
            Api.ClearCustomFunctions();
            
            // Display a message indicating that all custom functions were removed
            // Отображение сообщения о том, что все пользовательские функции были удалены
            worksheet.GetRange("A3").SetValue("All the custom functions were removed.");
            
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
    - id: "value-a3"
      type: "cell_value"
      sheet: "Active"
      cell: "A3"
      expected: "All the custom functions were removed."
      normalize: "string"
  hints:
    - "Начните с Api.GetActiveSheet()"
    - "Сверяйтесь с ожидаемым состоянием после запуска"
  source:
    category: "Api"
    examplePath: "modules/macros_ide/resources/learning/examples-macros/Api/ClearCustomFunctions.md"
  draft: true
---
# Теория
Черновая карточка создана автоматически на базе примера `Api/ClearCustomFunctions`.

Перед публикацией проверьте формулировки задания, подсказки и корректность проверок.
