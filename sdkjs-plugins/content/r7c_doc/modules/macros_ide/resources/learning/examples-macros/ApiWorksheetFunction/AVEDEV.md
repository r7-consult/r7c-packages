/**
 * @file AVEDEV_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.AVEDEV
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the average of the absolute deviations of data points from their mean.
 * It calculates the average deviation of a set of numbers and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть среднее значение абсолютных отклонений точек данных от их среднего значения.
 * Он вычисляет среднее отклонение набора чисел и отображает результат в ячейке A1.
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
        worksheet.GetRange("A1").SetValue(func.AVEDEV(78, 98, 123, 45, 70, 67, 3, 9, 289));
        
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
