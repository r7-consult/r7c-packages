/**
 * @file BIN2DEC_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.BIN2DEC
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to convert a binary number to decimal using ApiWorksheetFunction.BIN2DEC.
 * It converts the binary number 1110011100 to its decimal equivalent and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как преобразовать двоичное число в десятичное с помощью ApiWorksheetFunction.BIN2DEC.
 * Он преобразует двоичное число 1110011100 в его десятичный эквивалент и отображает результат в ячейке A1.
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
        // This example shows how to convert a binary number to decimal.
        
        // How to get a decimal representation of a binary number.
        
        // Use function to convert a binary to decimal.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.BIN2DEC(1110011100));
        
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
