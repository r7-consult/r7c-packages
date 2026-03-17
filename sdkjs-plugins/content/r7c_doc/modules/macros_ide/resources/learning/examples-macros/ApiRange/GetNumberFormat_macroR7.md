/**
 * @file GetNumberFormat_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetNumberFormat
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the number format of a range.
 * It sets a value in cell B2, gets its number format, and then displays it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить числовой формат диапазона.
 * Он устанавливает значение в ячейке B2, получает его числовой формат, а затем отображает его.
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
        // This example shows how to get a value that represents the format code for the current range.
        
        // How to find out a number format of a range.
        
        // Get a range, get its cell number format and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B2");
        range.SetValue(3);
        let format = range.GetNumberFormat();
        worksheet.GetRange("B3").SetValue("Number format: " + format);
        
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
