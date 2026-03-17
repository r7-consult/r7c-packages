/**
 * @file CreateNewHistoryPoint_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.CreateNewHistoryPoint
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to create a new history point in the document.
 * It sets a sample text in cell A1 of the active worksheet and then calls `Api.CreateNewHistoryPoint()`
 * to mark a new point in the document's history, displaying a confirmation message in cell A3.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как создать новую точку истории в документе.
 * Он устанавливает примерный текст в ячейке A1 активного листа, а затем вызывает `Api.CreateNewHistoryPoint()`
 * для отметки новой точки в истории документа, отображая сообщение о подтверждении в ячейке A3.
 *
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
        
        // Get the active worksheet
        // Получение активного листа
        var worksheet = Api.GetActiveSheet();
        
        // Set sample text in cell A1
        // Установка примерного текста в ячейке A1
        worksheet.GetRange("A1").SetValue("This is just a sample text.");
        
        // Create a new history point
        // Создание новой точки истории
        Api.CreateNewHistoryPoint();
        
        // Display a message confirming the creation of a new history point
        // Отображение сообщения, подтверждающего создание новой точки истории
        worksheet.GetRange("A3").SetValue("New history point was just created.");
        
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
