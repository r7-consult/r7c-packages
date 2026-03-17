/**
 * @file GetReferenceStyle_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetReferenceStyle
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve the current reference style of the spreadsheet.
 * It gets the active reference style (e.g., A1 or R1C1) and displays it in cell A1
 * of the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить текущий стиль ссылки электронной таблицы.
 * Он получает активный стиль ссылки (например, A1 или R1C1) и отображает его в ячейке A1
 * активного листа.
 *
 * @returns {string} The current reference style (e.g., "xlA1" or "xlR1C1"). (Текущий стиль ссылки.)
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
        
        // Get the current reference style and display it in cell A1
        // Получение текущего стиля ссылки и отображение его в ячейке A1
        worksheet.GetRange("A1").SetValue(Api.GetReferenceStyle());
        
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
