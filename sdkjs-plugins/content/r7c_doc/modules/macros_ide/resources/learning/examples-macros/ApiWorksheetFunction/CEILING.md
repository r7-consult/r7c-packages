/**
 * @file CEILING_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.CEILING
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to round a number up, to the nearest multiple of significance using ApiWorksheetFunction.CEILING.
 * It rounds 1.23 up to the nearest multiple of 0.1 and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как округлить число до ближайшего кратного значения с помощью ApiWorksheetFunction.CEILING.
 * Он округляет 1,23 до ближайшего кратного 0,1 и отображает результат в ячейке A1.
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
        // This example shows how to round a number up, to the nearest multiple of significance.
        
        // How to round a number up.
        
        // Use function to round a number up the nearest integer or to the nearest multiple of significance.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.CEILING(1.23, 0.1));
        
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
