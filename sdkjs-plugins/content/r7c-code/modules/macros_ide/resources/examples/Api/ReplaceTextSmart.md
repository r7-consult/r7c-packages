/**
 * @file ReplaceTextSmart_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.ReplaceTextSmart
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to intelligently replace text within a selected range in a spreadsheet.
 * It sets initial values in cells A1 and A2, selects the range A1:A2, and then replaces the text
 * in these cells with new values provided in an array.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как интеллектуально заменять текст в выделенном диапазоне в электронной таблице.
 * Он устанавливает начальные значения в ячейках A1 и A2, выбирает диапазон A1:A2, а затем заменяет текст
 * в этих ячейках новыми значениями, предоставленными в массиве.
 *
 * @param {Array<string>} newTexts - An array of strings to replace the existing text. (Массив строк для замены существующего текста.)
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
        
        // Set initial values in cells A1 and A2
        // Установка начальных значений в ячейках A1 и A2
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("A2").SetValue("2");
        
        // Select the range A1:A2
        // Выбор диапазона A1:A2
        let range = worksheet.GetRange("A1:A2");
        range.Select();
        
        // Replace text in the selected range with new values
        // Замена текста в выделенном диапазоне новыми значениями
        Api.ReplaceTextSmart(["Cell 1", "Cell 2"]);
        
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
