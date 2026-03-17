/**
 * @file GetTopMargin_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetTopMargin
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the top margin of the sheet.
 * It retrieves the top margin and displays it in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить верхнее поле листа.
 * Он извлекает верхнее поле и отображает его в ячейке A1.
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
        // This example shows how to get the top margin of the sheet.
        
        // How to get margin of the sheet's top side.
        
        // Get the size of the top margin of the sheet.
        
        let worksheet = Api.GetActiveSheet();
        let topMargin = worksheet.GetTopMargin();
        worksheet.GetRange("A1").SetValue("Top margin: " + topMargin + " mm");
        
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
