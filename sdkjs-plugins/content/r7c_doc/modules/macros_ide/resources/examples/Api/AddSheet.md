/**
 * @file AddSheet_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.AddSheet
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add a new worksheet to the current workbook.
 * It creates a new sheet named "New sheet".
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавить новый лист в текущую книгу.
 * Он создает новый лист с именем "New sheet".
 *
 * @param {string} sheetName - The name of the new sheet to add. (Имя нового листа для добавления.)
 * @returns {ApiWorksheet} The newly created ApiWorksheet object. (Вновь созданный объект ApiWorksheet.)
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
        
        // Add a new sheet named "New sheet"
        // Добавление нового листа с именем "New sheet"
        let sheet = Api.AddSheet("New sheet");
        
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
