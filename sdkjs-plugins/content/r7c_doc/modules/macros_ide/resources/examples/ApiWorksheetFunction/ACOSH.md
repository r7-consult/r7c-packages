/**
 * @file ACOSH_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.ACOSH
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the inverse hyperbolic cosine of a number.
 * It calculates the inverse hyperbolic cosine of 3.25 and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть обратный гиперболический косинус числа.
 * Он вычисляет обратный гиперболический косинус 3,25 и отображает результат в ячейке A1.
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
        // This example shows how to return the inverse hyperbolic cosine of a number.
        
        // How to get an inverse hyperbolic cosine of a number and display it in the worksheet.
        
        // Get a function that gets inverse hyperbolic cosine of a number.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.ACOSH(3.25));
        
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
