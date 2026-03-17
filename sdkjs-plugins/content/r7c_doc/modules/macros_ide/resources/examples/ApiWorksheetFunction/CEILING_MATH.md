/**
 * @file CEILING_MATH_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.CEILING_MATH
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to round a number up, to the nearest integer or to the nearest multiple of significance using ApiWorksheetFunction.CEILING_MATH.
 * It rounds -5.5 up to the nearest multiple of 2 with mode 1, and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как округлить число до ближайшего целого числа или до ближайшего кратного значения.
 * Он округляет -5,5 до ближайшего кратного 2 с режимом 1 и отображает результат в ячейке A1.
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
        // This example shows how to round a number up, to the nearest integer or to the nearest multiple of significance.
        
        // How to round a number up.
        
        // Use function to round a number up the nearest integer or to the nearest multiple of significance.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.CEILING_MATH(-5.5, 2, 1));
        
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
