/**
 * @file GetRows_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetRows
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the ApiRange object that represents all the cells on the rows range.
 * It gets the rows from range 1:4 and then sets their fill color.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект ApiRange, представляющий все ячейки в диапазоне строк.
 * Он получает строки из диапазона 1:4, а затем устанавливает их цвет заливки.
 *
 * @returns {void}
 *
 * @see https://r7-consult.com/
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
        // This example shows how to get the ApiRange object that represents all the cells on the rows range.
        
        // How to get all row cells.
        
        // Get all row cells from the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRows("1:4").SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        
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
