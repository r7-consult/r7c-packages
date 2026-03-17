/**
 * @file COSH_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.COSH
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the hyperbolic cosine of a number using ApiWorksheetFunction.COSH.
 * It calculates the hyperbolic cosine of 3 and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить гиперболический косинус числа с помощью ApiWorksheetFunction.COSH.
 * Он вычисляет гиперболический косинус 3 и отображает результат в ячейке A1.
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
        // This example shows how to get the hyperbolic cosine of a number.
        
        // How to find a hyperbolic cosine.
        
        // Use function to get the hyperbolic cosine of an angle.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.COSH(3));
        
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
