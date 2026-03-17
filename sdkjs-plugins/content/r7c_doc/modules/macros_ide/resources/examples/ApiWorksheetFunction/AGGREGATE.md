/**
 * @file AGGREGATE_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.AGGREGATE
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return an aggregate in a list or database.
 * It calculates the aggregate of numbers (10, 30, 50) using function 9 (SUM) and option 4 (ignore hidden rows, error values, nested SUBTOTAL and AGGREGATE functions), and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть агрегат в списке или базе данных.
 * Он вычисляет агрегат чисел (10, 30, 50) с использованием функции 9 (SUM) и опции 4 (игнорировать скрытые строки, значения ошибок, вложенные функции SUBTOTAL и AGGREGATE), а затем отображает результат в ячейке A1.
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
        // This example shows how to return an aggregate in a list or database.
        
        // How to get an aggregate of a numbers and display it in the worksheet.
        
        // Get a function that gets an aggregate from a list of numbers.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.AGGREGATE(9, 4, 10, 30, 50, 5));
        
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
