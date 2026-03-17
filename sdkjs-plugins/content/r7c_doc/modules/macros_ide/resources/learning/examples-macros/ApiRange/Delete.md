/**
 * @file Delete_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.Delete
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to delete a range from the worksheet.
 * It sets values in cells B4, C4, D4, and C5, and then deletes cell C4, shifting cells up.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как удалить диапазон из листа.
 * Он устанавливает значения в ячейках B4, C4, D4 и C5, а затем удаляет ячейку C4, сдвигая ячейки вверх.
 *
 * @returns {void}
 *
 * @see https://r7-consult.ru/
 */

(function() {
    'use strict';
    
    try {
        // Initialize R7 Office API
        const api = Api;
        if (!api) {
            throw new Error('R7 Office API not available');
        }
        
        // Original code enhanced with error handling:
        // This example deletes the Range object.
        
        // How to remove a range from the worksheet.
        
        // Get a range from the worksheet and delete it specifying the direction.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B4").SetValue("1");
        worksheet.GetRange("C4").SetValue("2");
        worksheet.GetRange("D4").SetValue("3");
        worksheet.GetRange("C5").SetValue("5");
        let range = worksheet.GetRange("C4");
        range.Delete("up");
        
        // Success notification
        console.log('Macro executed successfully');
        
    } catch (error) {
        console.error('Macro execution failed:', error.message);
        // Optional: Show error to user
        if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
            const sheet = Api.GetActiveSheet();
            if (sheet) {
                sheet.GetRange('A1').SetValue('Error: ' + error.message);
            }
        }
    }
})();
