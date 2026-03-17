/**
 * @file GetCommentById_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetCommentById
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve a specific comment from the document by its ID.
 * It first adds a comment, gets its ID, then retrieves the comment using that ID,
 * and displays the comment's text and author in cells A1 and B1 of the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить конкретный комментарий из документа по его ID.
 * Сначала он добавляет комментарий, получает его ID, затем извлекает комментарий по этому ID,
 * и отображает текст и автора комментария в ячейках A1 и B1 активного листа.
 *
 * @param {string} commentId - The ID of the comment to retrieve. (ID комментария для получения.)
 * @returns {ApiComment} The ApiComment object representing the retrieved comment. (Объект ApiComment, представляющий полученный комментарий.)
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
        
        // Add a comment and get its ID
        // Добавление комментария и получение его ID
        let comment = Api.AddComment("Comment", "Bob");
        let id = comment.GetId();
        
        // Retrieve the comment by its ID
        // Получение комментария по его ID
        comment = Api.GetCommentById(id);
        
        // Get the active worksheet
        // Получение активного листа
        let worksheet = Api.GetActiveSheet();
        
        // Display the comment's text and author in cells A1 and B1
        // Отображение текста и автора комментария в ячейках A1 и B1
        worksheet.GetRange("A1").SetValue("Comment Text: " + comment.GetText());
        worksheet.GetRange("B1").SetValue("Comment Author: " + comment.GetAuthorName());
        
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
