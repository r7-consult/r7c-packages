---
interactive:
  version: 1
  enabled: true
  lessonId: "interactive-draft-001"
  title: "Черновик: Api / AddComment"
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
            
            // Add two comments to the document
            // Добавление двух комментариев в документ
            Api.AddComment("Comment 1", "Bob");
            Api.AddComment("Comment 2");
            
            // Retrieve all comments from the document
            // Извлечение всех комментариев из документа
            let comments = Api.GetComments();
            
            // Get the active worksheet
            // Получение активного листа
            let worksheet = Api.GetActiveSheet();
            
            // Display the text of the first comment in cell A1
            // Отображение текста первого комментария в ячейке A1
            worksheet.GetRange("A1").SetValue("Comment Text: " + comments[0].GetText());
            
            // Display the author of the first comment in cell B1
            // Отображение автора первого комментария в ячейке B1
            worksheet.GetRange("B1").SetValue("Comment Author: " + comments[0].GetAuthorName());
            
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
    examplePath: "modules/macros_ide/resources/learning/examples-macros/Api/AddComment.md"
  draft: true
---
# Теория
Черновая карточка создана автоматически на базе примера `Api/AddComment`.

Перед публикацией проверьте формулировки задания, подсказки и корректность проверок.
