/**
 * @file BIN2OCT_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.BIN2OCT
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to convert a binary number to octal using ApiWorksheetFunction.BIN2OCT.
 * It converts the binary number 1110011100 to its octal equivalent with 4 characters, and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как преобразовать двоичное число в восьмеричное с помощью ApiWorksheetFunction.BIN2OCT.
 * Он преобразует двоичное число 1110011100 в его восьмеричный эквивалент с 4 символами и отображает результат в ячейке A1.
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
        // This example shows how to convert a binary number to octal.
        
        // How to get a octal representation of a binary number.
        
        // Use function to convert a binary to octal.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.BIN2OCT(1110011100, 4));
        
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
