/**
 * @file BESSELJ_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.BESSELJ
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the Bessel function Jn(x) using ApiWorksheetFunction.BESSELJ.
 * It calculates the Bessel function for x=1.9 and n=2, and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть функцию Бесселя Jn(x) с помощью ApiWorksheetFunction.BESSELJ.
 * Он вычисляет функцию Бесселя для x=1,9 и n=2 и отображает результат в ячейке A1.
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
        // This example shows how to return the Bessel function Jn(x).
        
        // How to get a result from Bessel function Jn(x).
        
        // Use function to get the Bessel function Jn(x).
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.BESSELJ(1.9, 2));
        
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
