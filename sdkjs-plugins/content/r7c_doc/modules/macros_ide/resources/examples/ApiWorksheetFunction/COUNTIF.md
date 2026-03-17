/**
 * @file COUNTIF_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.COUNTIF
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to count a number of cells within a range that meet the given condition using ApiWorksheetFunction.COUNTIF.
 * It sets up lists of fruits and numbers, and then counts the number of cells in the range A1:B3 that end with "es", displaying the result in cell D3.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как подсчитать количество ячеек в диапазоне, которые соответствуют заданному условию, с помощью ApiWorksheetFunction.COUNTIF.
 * Он настраивает списки фруктов и чисел, а затем подсчитывает количество ячеек в диапазоне A1:B3, которые заканчиваются на «es», отображая результат в ячейке D3.
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
        // This example shows how to count a number of cells within a range that meet the given condition.
        
        // How to find a number of cells that satisfies some condition.
        
        // Use function to get cells if a condition is met.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let fruits = ["Apples", "ranges", "Bananas"];
        let numbers = [45, 6, 8];
        
        for (let i = 0; i < fruits.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(fruits[i]);
        }
        for (let j = 0; j < numbers.length; j++) {
            worksheet.GetRange("B" + (j + 1)).SetValue(numbers[j]);
        }
        
        let range = worksheet.GetRange("A1:B3");
        worksheet.GetRange("D3").SetValue(func.COUNTIF(range, "*es"));
        
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
