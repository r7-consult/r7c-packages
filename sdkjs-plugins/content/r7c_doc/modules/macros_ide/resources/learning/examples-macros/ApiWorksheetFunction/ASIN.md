/**
 * @file ASIN_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.ASIN
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the arcsine of a number in radians, in the range from -Pi/2 to Pi/2.
 * It calculates the arcsine of 0.25 and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть арксинус числа в радианах в диапазоне от -Pi/2 до Pi/2.
 * Он вычисляет арксинус 0,25 и отображает результат в ячейке A1.
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
        // This example shows how to return the arcsine of a number in radians, in the range from Pi/2 to Pi/2.
        
        // How to get an arcsine of a number in radians.
        
        // Use function to get an arcsine of a number and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.ASIN(0.25));
        
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
