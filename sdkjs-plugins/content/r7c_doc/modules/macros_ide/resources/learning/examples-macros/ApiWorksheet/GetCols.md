/**
 * @file GetCols_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetCols
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the ApiRange object that represents all the cells on the columns range.
 * It gets the columns from range A1:C1 and then sets their fill color.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект ApiRange, представляющий все ячейки в диапазоне столбцов.
 * Он получает столбцы из диапазона A1:C1, а затем устанавливает их цвет заливки.
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
        // This example shows how to get the ApiRange object that represents all the cells on the columns range.
        
        // How to get all column cells.
        
        // Get all column cells from the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let cols = worksheet.GetCols("A1:C1");
        cols.SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        
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
