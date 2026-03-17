/**
 * @file GetSheet_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetSheet
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve a specific sheet from the workbook by its name.
 * It gets the sheet named "Sheet1" and then sets a sample text in cell A1 of that sheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить определенный лист из книги по его имени.
 * Он получает лист с именем "Sheet1", а затем устанавливает примерный текст в ячейке A1 этого листа.
 *
 * @param {string} sheetName - The name of the sheet to retrieve. (Имя листа для получения.)
 * @returns {ApiWorksheet} The ApiWorksheet object representing the retrieved sheet. (Объект ApiWorksheet, представляющий полученный лист.)
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
        
        // Get the sheet named "Sheet1"
        // Получение листа с именем "Sheet1"
        let worksheet = Api.GetSheet("Sheet1");
        
        // Set a sample text in cell A1 of the retrieved sheet
        // Установка примерного текста в ячейке A1 полученного листа
        worksheet.GetRange("A1").SetValue("This is a sample text on 'Sheet1'.");
        
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
