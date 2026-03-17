/**
 * @file ABS_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.ABS
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the absolute value of a number using ApiWorksheetFunction.ABS.
 * It retrieves the absolute value of -123.14 and displays it in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить абсолютное значение числа с помощью ApiWorksheetFunction.ABS.
 * Он извлекает абсолютное значение -123,14 и отображает его в ячейке A1.
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
        // This example shows how to get absolute value of a number.
        
        // How to add absolute value to the worksheet.
        
        // Get a function that gets absolute value.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.ABS(-123.14));
        
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
