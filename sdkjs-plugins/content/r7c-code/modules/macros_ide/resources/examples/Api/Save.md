/**
 * @file Save_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.Save
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to save changes to the current document.
 * It sets a sample text in cell A1 of the active worksheet and then calls the `Api.Save()` method
 * to persist these changes.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как сохранить изменения в текущем документе.
 * Он устанавливает примерный текст в ячейке A1 активного листа, а затем вызывает метод `Api.Save()`
 * для сохранения этих изменений.
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
        let worksheet = Api.GetActiveSheet();
        
        // Set a sample text in cell A1
        // Установка примерного текста в ячейке A1
        worksheet.GetRange("A1").SetValue("This sample text is saved to the worksheet.");
        
        // Save the document
        // Сохранение документа
        Api.Save();
        
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
