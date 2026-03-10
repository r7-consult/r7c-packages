/**
 * @file ACOT_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.ACOT
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the arccotangent of a number, in radians in the range from 0 to Pi.
 * It calculates the arccotangent of 0 and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть арккотангенс числа в радианах в диапазоне от 0 до Pi.
 * Он вычисляет арккотангенс 0 и отображает результат в ячейке A1.
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
        // This example shows how to return the arccotangent of a number, in radians in the range from 0 to Pi.
        
        // How to get an arccotangent of a number and display it in the worksheet.
        
        // Get a function that gets arccotangent of a number.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.ACOT(0));
        
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
