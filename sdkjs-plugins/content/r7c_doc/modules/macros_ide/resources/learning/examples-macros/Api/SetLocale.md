/**
 * @file SetLocale_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.SetLocale
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the locale for the document.
 * It sets the document's locale to "en-CA" (English - Canada) and displays a message
 * in cell A1 of the active worksheet indicating the change.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить локаль для документа.
 * Он устанавливает локаль документа на "en-CA" (английский - Канада) и отображает сообщение
 * в ячейке A1 активного листа, указывающее на изменение.
 *
 * @param {string} locale - The locale string to set (e.g., "en-US", "ru-RU", "en-CA"). (Строка локали для установки.)
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
        
        // Set the document locale to English (Canada)
        // Установка локали документа на английский (Канада)
        Api.SetLocale("en-CA");
        
        // Display a message indicating the locale change
        // Отображение сообщения, указывающего на изменение локали
        worksheet.GetRange("A1").SetValue("A sample spreadsheet with the language set to English (Canada).");
        
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
