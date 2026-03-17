/**
 * @file CEILING_PRECISE_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.CEILING_PRECISE
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return a number that is rounded up to the nearest integer or to the nearest multiple of significance, regardless of its sign, using ApiWorksheetFunction.CEILING_PRECISE.
 * It rounds -6.7 up to the nearest multiple of 2 and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть число, округленное до ближайшего целого числа или до ближайшего кратного значения, независимо от его знака, с помощью ApiWorksheetFunction.CEILING_PRECISE.
 * Он округляет -6,7 до ближайшего кратного 2 и отображает результат в ячейке A1.
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
        // This example shows how to return a number that is rounded up to the nearest integer or to the nearest multiple of significance. The number is always rounded up regardless of its sing.
        
        // How to round a number up precisely.
        
        // Use function to round a negative or positive number up the nearest integer or to the nearest multiple of significance.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.CEILING_PRECISE(-6.7, 2));
        
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
