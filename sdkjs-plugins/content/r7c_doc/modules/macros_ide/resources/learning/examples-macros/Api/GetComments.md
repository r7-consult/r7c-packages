/**
 * @file GetComments_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetComments
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve an array of `ApiComment` objects from the document.
 * It adds two comments, then retrieves all comments, and displays the text and author
 * of the first comment in cells A1 and B1 of the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить массив объектов `ApiComment` из документа.
 * Он добавляет два комментария, затем извлекает все комментарии и отображает текст и автора
 * первого комментария в ячейках A1 и B1 активного листа.
 *
 * @returns {Array<ApiComment>} An array of ApiComment objects representing all comments in the document. (Массив объектов ApiComment, представляющих все комментарии в документе.)
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
        
        // Add two comments for demonstration
        // Добавление двух комментариев для демонстрации
        Api.AddComment("Comment 1", "Bob");
        Api.AddComment("Comment 2", "Bob");
        
        // Get all comments in the document
        // Получение всех комментариев в документе
        let arrComments = Api.GetComments();
        
        // Get the active worksheet
        // Получение активного листа
        let worksheet = Api.GetActiveSheet();
        
        // Display the text and author of the first comment in cells A1 and B1
        // Отображение текста и автора первого комментария в ячейках A1 и B1
        worksheet.GetRange("A1").SetValue("Comment Text: " + arrComments[0].GetText());
        worksheet.GetRange("B1").SetValue("Comment Author: " + arrComments[0].GetAuthorName());
        
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
