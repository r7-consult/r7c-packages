/**
 * @file ARABIC_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.ARABIC
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to convert a Roman numeral to Arabic using ApiWorksheetFunction.ARABIC.
 * It converts "MCCL" to its Arabic equivalent and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как преобразовать римскую цифру в арабскую с помощью ApiWorksheetFunction.ARABIC.
 * Он преобразует «MCCL» в его арабский эквивалент и отображает результат в ячейке A1.
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
        // This example shows how to convert a Roman numeral to Arabic.
        
        // How to convert numbers to Arabic numerical.
        
        // Use function to convert numbers to Arabic numerical.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.ARABIC("MCCL"));
        
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
