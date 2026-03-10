/**
 * @file AVERAGE_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.AVERAGE
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the average of a set of numbers.
 * It calculates the average of a given set of numbers and displays the result in cell B2.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть среднее значение набора чисел.
 * Он вычисляет среднее значение заданного набора чисел и отображает результат в ячейке B2.
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
        // This example shows how to return the average of the absolute deviations of data points from their mean.
        
        // How to get an average of the absolute deviations.
        
        // Use function to get the average of the absolute deviations of data points from their mean.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let ans = func.AVERAGE(123, 197, 46, 345, 67, 456);
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
