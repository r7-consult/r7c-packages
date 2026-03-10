/**
 * @file Format_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.Format
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to format a value according to a specified format expression
 * using `Api.Format()`.
 * It formats the number "123456" as currency "$#,##0" and displays the formatted result
 * in cell A1 of the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как форматировать значение в соответствии с заданным выражением формата
 * с помощью `Api.Format()`.
 * Он форматирует число "123456" как валюту "$#,##0" и отображает отформатированный результат
 * в ячейке A1 активного листа.
 *
 * @param {any} value - The value to be formatted. (Значение для форматирования.)
 * @param {string} formatExpression - The format expression string (e.g., "$#,##0", "yyyy-mm-dd"). (Строка выражения формата.)
 * @returns {string} The formatted value as a string. (Отформатированное значение в виде строки.)
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
        
        // Format the number and set the result in cell A1
        // Форматирование числа и установка результата в ячейке A1
        let format = Api.Format("123456", "$#,##0");
        worksheet.GetRange("A1").SetValue(format);
        
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
