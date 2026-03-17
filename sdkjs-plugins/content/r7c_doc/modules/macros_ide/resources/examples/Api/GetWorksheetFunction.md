/**
 * @file GetWorksheetFunction_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetWorksheetFunction
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to use built-in worksheet functions through the `Api.GetWorksheetFunction()` method.
 * It retrieves the worksheet function object and then uses the `ASC` function to convert full-width
 * (double-byte) characters to half-width (single-byte) characters in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как использовать встроенные функции листа через метод `Api.GetWorksheetFunction()`.
 * Он извлекает объект функции листа, а затем использует функцию `ASC` для преобразования полноширинных
 * (двухбайтовых) символов в полуширинные (однобайтовые) символы в ячейке A1.
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
        
        // Get the worksheet function object
        // Получение объекта функции листа
        let func = Api.GetWorksheetFunction();
        
        // Use the ASC function and set the result in cell A1
        // Использование функции ASC и установка результата в ячейке A1
        worksheet.GetRange("A1").SetValue(func.ASC("text"));
        
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
