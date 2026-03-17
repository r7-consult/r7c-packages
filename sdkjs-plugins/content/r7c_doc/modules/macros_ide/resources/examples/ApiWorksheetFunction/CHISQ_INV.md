/**
 * @file CHISQ_INV_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.CHISQ_INV
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the inverse of the left-tailed probability of the chi-squared distribution.
 * It calculates the inverse of the left-tailed probability for specified parameters and displays the result in cell B2.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть обратную левостороннюю вероятность распределения хи-квадрат.
 * Он вычисляет обратную левостороннюю вероятность для указанных параметров и отображает результат в ячейке B2.
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
        // This example shows how to return the inverse of the left-tailed probability of the chi-squared distribution.
        
        // How to return the inverse of the inverse of the left-tailed probability of the chi-squared distribution.
        
        // Use function to return the inverse of the left-tailed probability of the chi-squared distribution.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let ans = func.CHISQ_INV(0.4, 10);
        worksheet.GetRange("B2").SetValue(ans);
        
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
