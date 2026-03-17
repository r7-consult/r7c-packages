/**
 * @file BITRSHIFT_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.BITRSHIFT
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return a number shifted right by the specified number of bits using ApiWorksheetFunction.BITRSHIFT.
 * It calculates the bitwise right shift of 13 by 2 bits and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть число, сдвинутое вправо на указанное количество бит с помощью ApiWorksheetFunction.BITRSHIFT.
 * Он вычисляет побитовый сдвиг вправо числа 13 на 2 бита и отображает результат в ячейке A1.
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
        // This example shows how to return a number shifted right by the specified number of bits. 
        
        // How to get a result from bits right shift.
        
        // Use function to calculate bitwise right shift operation.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.BITRSHIFT(13, 2));
        
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
