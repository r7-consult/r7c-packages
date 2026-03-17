/**
 * @file COMPLEX_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.COMPLEX
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to convert real and imaginary coefficients into a complex number using ApiWorksheetFunction.COMPLEX.
 * It converts real part -2 and imaginary part 2.5 with suffix "i" into a complex number and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как преобразовать действительные и мнимые коэффициенты в комплексное число с помощью ApiWorksheetFunction.COMPLEX.
 * Он преобразует действительную часть -2 и мнимую часть 2,5 с суффиксом «i» в комплексное число и отображает результат в ячейке A1.
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
        // This example shows how to convert real and imaginary coefficients into a complex number.
        
        // How to create a complex number using coefficients.
        
        // Use function to convert real and imaginary coefficients into a complex number.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.COMPLEX(-2, 2.5, "i"));
        
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
