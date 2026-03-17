/**
 * @file GetCol_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetCol
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the column number for the selected cell.
 * It gets the column of cell D9 and then displays its value.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить номер столбца для выбранной ячейки.
 * Он получает столбец ячейки D9, а затем отображает его значение.
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
        // This example shows how to get a column number for the selected cell.
        
        // How to get a cell column index.
        
        // Get a range and display its column number.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("D9").GetCol();
        worksheet.GetRange("A2").SetValue(range.toString());
        
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
