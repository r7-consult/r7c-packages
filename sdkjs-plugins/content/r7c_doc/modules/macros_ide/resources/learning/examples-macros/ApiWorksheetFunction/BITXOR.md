/**
 * @file BITXOR_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.BITXOR
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return a bitwise "XOR" of two numbers using ApiWorksheetFunction.BITXOR.
 * It calculates the bitwise XOR of 5 and 3 and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть побитовое «XOR» двух чисел с помощью ApiWorksheetFunction.BITXOR.
 * Он вычисляет побитовое XOR для 5 и 3 и отображает результат в ячейке A1.
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
        // This example shows how to return a bitwise "XOR" of two numbers. 
        
        // How to get a result from XOR operation.
        
        // Use function to calculate bitwise "XOR" operation.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.BITXOR(5, 3));
        
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
