/**
 * @file BITAND_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.BITAND
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return a bitwise "AND" of two numbers using ApiWorksheetFunction.BITAND.
 * It calculates the bitwise AND of 13 and 25, and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть побитовое «И» двух чисел с помощью ApiWorksheetFunction.BITAND.
 * Он вычисляет побитовое И для 13 и 25 и отображает результат в ячейке A1.
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
        // This example shows how to return a bitwise "AND" of two numbers. 
        
        // How to get a two numbers bitwise AND.
        
        // Use function to calculate bitwise AND opertation.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.BITAND(13, 25));
        
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
