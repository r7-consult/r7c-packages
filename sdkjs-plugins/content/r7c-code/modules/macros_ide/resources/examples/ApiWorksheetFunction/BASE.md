/**
 * @file BASE_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.BASE
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to convert a number into a text representation with the given radix (base) using ApiWorksheetFunction.BASE.
 * It converts the number 5 to base 2 with a minimum length of 5, and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как преобразовать число в текстовое представление с заданным основанием (базисом) с помощью ApiWorksheetFunction.BASE.
 * Он преобразует число 5 в основание 2 с минимальной длиной 5 и отображает результат в ячейке A1.
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
        // This example shows how to convert a number into a text representation with the given radix (base).
        
        // How to convert a number into text.
        
        // Use function to get a text from a number.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.BASE(5, 2, 5));
        
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
