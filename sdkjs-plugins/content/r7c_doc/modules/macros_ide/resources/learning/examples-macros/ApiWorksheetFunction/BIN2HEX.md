/**
 * @file BIN2HEX_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.BIN2HEX
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to convert a binary number to hexadecimal using ApiWorksheetFunction.BIN2HEX.
 * It converts the binary number 1110011100 to its hexadecimal equivalent with 4 characters, and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как преобразовать двоичное число в шестнадцатеричное с помощью ApiWorksheetFunction.BIN2HEX.
 * Он преобразует двоичное число 1110011100 в его шестнадцатеричный эквивалент с 4 символами и отображает результат в ячейке A1.
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
        // This example shows how to convert a binary number to hexadecimal.
        
        // How to get a hexadecimal representation of a binary number.
        
        // Use function to convert a binary to hexadecimal.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.BIN2HEX(1110011100, 4));
        
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
