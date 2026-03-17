/**
 * @file AVERAGEIF_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.AVERAGEIF
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to find the average (arithmetic mean) for the cells specified by a given condition or criteria.
 * It sets up a list of numbers in column A, calculates the average of numbers greater than 45, and displays the result in cell C1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как найти среднее (арифметическое) для ячеек, указанных заданным условием или критерием.
 * Он настраивает список чисел в столбце A, вычисляет среднее значение чисел, превышающих 45, и отображает результат в ячейке C1.
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
        // This example shows how to find the average (arithmetic mean) for the cells specified by a given condition or criteria.
        
        // How to find an average (arithmetic mean) using condition.
        
        // Use function to get an average of the cells if the condition is met.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let numbers = [67, 87, 98, 45];
        
        for (let i = 0; i < numbers.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(numbers[i]);
        }
        
        let range = worksheet.GetRange("A1:A4");
        worksheet.GetRange("C1").SetValue(func.AVERAGEIF(range, ">45"));
        
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
