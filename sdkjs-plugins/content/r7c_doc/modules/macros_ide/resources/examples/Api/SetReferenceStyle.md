/**
 * @file SetReferenceStyle_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.SetReferenceStyle
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the reference style for the spreadsheet.
 * It sets the reference style to "xlR1C1" (R1C1 reference style) and then displays
 * the currently active reference style in cell A1 of the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить стиль ссылки для электронной таблицы.
 * Он устанавливает стиль ссылки на "xlR1C1" (стиль ссылки R1C1), а затем отображает
 * текущий активный стиль ссылки в ячейке A1 активного листа.
 *
 * @param {string} style - The reference style to set (e.g., "xlA1" or "xlR1C1"). (Стиль ссылки для установки.)
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
        
        // Set the reference style to R1C1
        // Установка стиля ссылки на R1C1
        Api.SetReferenceStyle("xlR1C1");
        
        // Display the current reference style in cell A1
        // Отображение текущего стиля ссылки в ячейке A1
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
