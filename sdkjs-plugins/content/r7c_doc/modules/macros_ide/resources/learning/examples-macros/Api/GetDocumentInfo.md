/**
 * @file GetDocumentInfo_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetDocumentInfo
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve document information using `Api.GetDocumentInfo()`.
 * It gets the document information object and then displays the application name
 * (e.g., "Spreadsheet Editor") in cell A1 of the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить информацию о документе с помощью `Api.GetDocumentInfo()`.
 * Он получает объект информации о документе, а затем отображает имя приложения
 * (например, "Spreadsheet Editor") в ячейке A1 активного листа.
 *
 * @returns {object} An object containing various properties about the document, such as Application, Author, etc. (Объект, содержащий различные свойства документа, такие как Application, Author и т.д.)
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
        
        // Get document information
        // Получение информации о документе
        let docInfo = Api.GetDocumentInfo();
        
        // Get the active worksheet and range A1
        // Получение активного листа и диапазона A1
        let range = Api.GetActiveSheet().GetRange('A1');
        
        // Display the application name in cell A1
        // Отображение имени приложения в ячейке A1
        range.SetValue('This document has been created with: ' + docInfo.Application);
        
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
