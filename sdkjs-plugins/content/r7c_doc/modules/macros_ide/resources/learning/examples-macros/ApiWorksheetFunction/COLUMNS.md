/**
 * @file COLUMNS_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.COLUMNS
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the number of columns in the cell range using ApiWorksheetFunction.COLUMNS.
 * It sets up two columns of data, and then displays the number of columns in the range A1:B3 in cell B4.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть количество столбцов в диапазоне ячеек с помощью ApiWorksheetFunction.COLUMNS.
 * Он настраивает два столбца данных, а затем отображает количество столбцов в диапазоне A1:B3 в ячейке B4.
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
        // This example shows how to return the number of columns in the cell range.
        
        // How to find a number of columns from a range.
        
        // Use function to count range column.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let column1 = [13, 14, 15];
        let column2 = [23, 24, 25];
        
        for (let i = 0; i < column1.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(column1[i]);
        }
        for (let j = 0; j < column2.length; j++) {
            worksheet.GetRange("B" + (j + 1)).SetValue(column2[j]);
        }
        
        let range = worksheet.GetRange("A1:B3");
        worksheet.GetRange("B4").SetValue(func.COLUMNS(range));
        
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
