/**
 * @file SetFormulaArray_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetFormulaArray
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the array formula of a range.
 * It sets an array formula in range A1:C3.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить формулу массива диапазона.
 * Он устанавливает формулу массива в диапазоне A1:C3.
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
        // This example sets the array formula of a range.
        
        // How to set the array formula value.
        
        // Set the array formula.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1:C3").SetFormulaArray("={1,2,3}");
        
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
