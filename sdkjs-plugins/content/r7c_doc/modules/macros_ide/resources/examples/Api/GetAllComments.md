/**
 * @file GetAllComments_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetAllComments
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve all comments from the worksheet.
 * It adds two comments (one directly via Api, one to a specific cell),
 * then retrieves all comments, and displays the text and author of the second comment
 * in cells A1 and A2 of the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить все комментарии с листа.
 * Он добавляет два комментария (один напрямую через Api, один к определенной ячейке),
 * затем извлекает все комментарии и отображает текст и автора второго комментария
 * в ячейках A1 и A2 активного листа.
 *
 * @returns {Array<ApiComment>} An array of ApiComment objects representing all comments in the worksheet. (Массив объектов ApiComment, представляющих все комментарии на листе.)
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
        
        // Get the active worksheet
        // Получение активного листа
        let worksheet = Api.GetActiveSheet();
        
        // Add comments for demonstration
        // Добавление комментариев для демонстрации
        Api.AddComment("Comment 1", "John Smith");
        worksheet.GetRange("A4").AddComment("Comment 2", "Mark Potato");
        
        // Get all comments in the document
        // Получение всех комментариев в документе
        let arrComments = Api.GetAllComments();
        
        // Display the text and author of the second comment in cells A1 and A2
        // Отображение текста и автора второго комментария в ячейках A1 и A2
        worksheet.GetRange("A1").SetValue("Comment text: " + arrComments[1].GetText());
        worksheet.GetRange("A2").SetValue("Comment author: " + arrComments[1].GetAuthorName());
        
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
