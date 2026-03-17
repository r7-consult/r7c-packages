/**
 * @file GetActiveSheet_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetActiveSheet
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve the currently active sheet in the workbook.
 * It gets the active worksheet and then performs a simple calculation by setting values
 * in cells B1 and B2, and a formula in B3 to multiply them, displaying the result.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить текущий активный лист в книге.
 * Он получает активный лист, а затем выполняет простой расчет, устанавливая значения
 * в ячейках B1 и B2, и формулу в B3 для их умножения, отображая результат.
 *
 * @returns {ApiWorksheet} The ApiWorksheet object representing the active sheet. (Объект ApiWorksheet, представляющий активный лист.)
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
        
        // Set values in cells B1 and B2
        // Установка значений в ячейках B1 и B2
        worksheet.GetRange("B1").SetValue("2");
        worksheet.GetRange("B2").SetValue("2");
        
        // Set a label in cell A3
        // Установка метки в ячейке A3
        worksheet.GetRange("A3").SetValue("2x2=");
        
        // Set a formula in cell B3 to multiply B1 and B2
        // Установка формулы в ячейке B3 для умножения B1 и B2
        worksheet.GetRange("B3").SetValue("=B1*B2");
        
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
