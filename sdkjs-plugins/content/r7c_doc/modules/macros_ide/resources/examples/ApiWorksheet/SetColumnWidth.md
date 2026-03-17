/**
 * @file SetColumnWidth_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.SetColumnWidth
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the width of the specified column.
 * It sets the width of column 0 to 10 and column 1 to 20.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить ширину указанного столбца.
 * Он устанавливает ширину столбца 0 на 10, а столбца 1 на 20.
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
        // This example sets the width of the specified column.
        
        // How to set a column width.
        
        // Resize column width.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetColumnWidth(0, 10);
        worksheet.SetColumnWidth(1, 20);
        
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
