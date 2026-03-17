/**
 * @file COUNTBLANK_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.COUNTBLANK
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to count a number of empty cells in a specified range of cells using ApiWorksheetFunction.COUNTBLANK.
 * It sets up numbers and strings in different columns, and then counts the number of empty cells in the range A1:C3, displaying the result in cell D3.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как подсчитать количество пустых ячеек в указанном диапазоне ячеек с помощью ApiWorksheetFunction.COUNTBLANK.
 * Он настраивает числа и строки в разных столбцах, а затем подсчитывает количество пустых ячеек в диапазоне A1:C3, отображая результат в ячейке D3.
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
        // This example shows how to counts a number of empty cells in a specified range of cells.
        
        // How to find a number of empty cells.
        
        // Use function to get empty cells count.
        
        let worksheet = Api.GetActiveSheet();
        let numbersArr = [45, 6, 8];
        let stringsArr = ["Apples", "ranges", "Bananas"]
        
        // Place the numbers in cells
        for (let i = 0; i < numbersArr.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(numbersArr[i]);
        }
        
        // Place the strings in cells
        for (let n = 0; n < stringsArr.length; n++) {
            worksheet.GetRange("B" + (n + 1)).SetValue(stringsArr[n]);
        }
        
        let func = Api.GetWorksheetFunction();
        let ans = func.COUNTBLANK(worksheet.GetRange("A1:C3"));
        worksheet.GetRange("D3").SetValue(ans);
        
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
