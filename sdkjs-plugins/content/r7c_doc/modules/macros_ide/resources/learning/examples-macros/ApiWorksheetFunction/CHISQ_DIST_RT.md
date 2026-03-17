/**
 * @file CHISQ_DIST_RT_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.CHISQ_DIST_RT
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the right-tailed probability of the chi-squared distribution.
 * It calculates the right-tailed probability for specified parameters and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть правостороннюю вероятность распределения хи-квадрат.
 * Он вычисляет правостороннюю вероятность для указанных параметров и отображает результат в ячейке A1.
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
        // This example shows how to return the right-tailed probability of the chi-squared distribution.
        
        // How to return the right-tailed probability of the chi-squared distribution.
        
        // Use function to return the right-tailed probability of the chi-squared distribution.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.CHISQ_DIST_RT(12, 10));
        
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
