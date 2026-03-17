/**
 * @file BESSELK_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.BESSELK
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the modified Bessel function Kn(x) using ApiWorksheetFunction.BESSELK.
 * It calculates the modified Bessel function for x=1.5 and n=1, and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть модифицированную функцию Бесселя Kn(x) с помощью ApiWorksheetFunction.BESSELK.
 * Он вычисляет модифицированную функцию Бесселя для x=1,5 и n=1 и отображает результат в ячейке A1.
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
        // This example shows how to return the modified Bessel function Kn(x).
        
        // How to get a result from Bessel function Kn(x).
        
        // Use function to get the Bessel function Kn(x).
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.BESSELK(1.5, 1));
        
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
