/**
 * @file AddComment_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.AddComment
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates the usage of the `Api.AddComment` method to add comments to a spreadsheet.
 * It adds two comments, retrieves all comments, and then displays the text and author of the first comment
 * in cells A1 and B1 of the active worksheet. This macro serves as an example for interacting with comments
 * within the R7 Office suite, showcasing how to programmatically add, retrieve, and display comment information.
 *
 * @description (Russian)
 * Этот макрос демонстрирует использование метода `Api.AddComment` для добавления комментариев в электронную таблицу.
 * Он добавляет два комментария, извлекает все комментарии, а затем отображает текст и автора первого комментария
 * в ячейках A1 и B1 активного листа. Этот макрос служит примером для взаимодействия с комментариями
 * в пакете R7 Office, демонстрируя, как программно добавлять, извлекать и отображать информацию о комментариях.
 *
 * @param {string} text - The text content of the comment. (Текстовое содержимое комментария.)
 * @param {string} author - The author's name for the comment. (Имя автора комментария.)
 * @returns {void}
 *
 * @see https://r7-consult.ru/
 */

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
