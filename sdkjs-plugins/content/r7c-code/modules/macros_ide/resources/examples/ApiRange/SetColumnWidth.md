/**
 * @file SetColumnWidth_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetColumnWidth
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the width of all columns in a range.
 * It sets the width of column A to 20.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить ширину всех столбцов в диапазоне.
 * Он устанавливает ширину столбца A на 20.
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
        // This example sets the width of all the columns in the range.
        
        // How to make a cell column wider.
        
        // Get a range and set its column width.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetColumnWidth(20);
        
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
