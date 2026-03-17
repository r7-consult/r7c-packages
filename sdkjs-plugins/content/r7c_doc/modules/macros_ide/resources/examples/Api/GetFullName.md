/**
 * @file GetFullName_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetFullName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve the full name of the currently opened file.
 * It gets the full name of the document and displays it in cell B1 of the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить полное имя текущего открытого файла.
 * Он получает полное имя документа и отображает его в ячейке B1 активного листа.
 *
 * @returns {string} The full name of the currently opened file. (Полное имя текущего открытого файла.)
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
        
        // Get the full name of the current document
        // Получение полного имени текущего документа
        let name = Api.GetFullName();
        
        // Display the full name in cell B1
        // Отображение полного имени в ячейке B1
        worksheet.GetRange("B1").SetValue("File name: " + name);
        
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
