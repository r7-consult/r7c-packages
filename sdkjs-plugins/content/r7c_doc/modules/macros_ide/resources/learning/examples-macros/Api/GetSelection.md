/**
 * @file GetSelection_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetSelection
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve the currently selected range in the spreadsheet.
 * It gets the selected range and then sets its value to "selected".
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить текущий выделенный диапазон в электронной таблице.
 * Он получает выделенный диапазон, а затем устанавливает его значение в "selected".
 *
 * @returns {ApiRange} The ApiRange object representing the selected range. (Объект ApiRange, представляющий выделенный диапазон.)
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
        
        // Get the currently selected range and set its value
        // Получение текущего выделенного диапазона и установка его значения
        Api.GetSelection().SetValue("selected");
        
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
