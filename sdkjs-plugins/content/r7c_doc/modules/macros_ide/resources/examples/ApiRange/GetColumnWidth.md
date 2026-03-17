/**
 * @file GetColumnWidth_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetColumnWidth
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the column width value of a range.
 * It gets the column width of cell A1 and then displays it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить значение ширины столбца диапазона.
 * Он получает ширину столбца ячейки A1, а затем отображает ее.
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
        // This example shows how to get the column width value.
        
        // How to get width of a range column.
        
        // Get a range, get its column width and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let width = worksheet.GetRange("A1").GetColumnWidth();
        worksheet.GetRange("A1").SetValue("Width: ");
        worksheet.GetRange("B1").SetValue(width);
        
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
