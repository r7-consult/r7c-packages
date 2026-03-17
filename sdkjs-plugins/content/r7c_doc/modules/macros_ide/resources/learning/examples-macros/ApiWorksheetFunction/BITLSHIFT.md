/**
 * @file BITLSHIFT_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.BITLSHIFT
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return a number shifted left by the specified number of bits using ApiWorksheetFunction.BITLSHIFT.
 * It calculates the bitwise left shift of 4 by 2 bits and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть число, сдвинутое влево на указанное количество бит с помощью ApiWorksheetFunction.BITLSHIFT.
 * Он вычисляет побитовый сдвиг влево числа 4 на 2 бита и отображает результат в ячейке A1.
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
        // This example shows how to return a number shifted left by the specified number of bits. 
        
        // How to get a result from bits left shift.
        
        // Use function to calculate bitwise left shift operation.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.BITLSHIFT(4, 2));
        
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
