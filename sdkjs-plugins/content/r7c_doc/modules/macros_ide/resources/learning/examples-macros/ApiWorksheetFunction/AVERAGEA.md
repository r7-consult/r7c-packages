/**
 * @file AVERAGEA_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.AVERAGEA
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the average (arithmetic mean) of the specified arguments.
 * It calculates the average of a given set of numbers and boolean values, and displays the result in cell B2.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть среднее (арифметическое) значение указанных аргументов.
 * Он вычисляет среднее значение заданного набора чисел и логических значений и отображает результат в ячейке B2.
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
        // This example shows how to return the average (arithmetic mean) of the specified arguments.
        
        // How to find an average (arithmetic mean).
        
        // Use function to get the find an average (arithmetic mean).
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let ans = func.AVERAGEA(78, 98, 46, 123, 45, true, false);
        worksheet.GetRange("B2").SetValue(ans);
        
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
