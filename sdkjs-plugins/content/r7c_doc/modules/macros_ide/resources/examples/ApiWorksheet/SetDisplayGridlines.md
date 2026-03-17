/**
 * @file SetDisplayGridlines_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.SetDisplayGridlines
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to specify whether the sheet gridlines must be displayed or not.
 * It sets a value in cell A2 and then sets the display gridlines to false.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как указать, должны ли отображаться линии сетки листа.
 * Он устанавливает значение в ячейке A2, а затем устанавливает отображение линий сетки в значение false.
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
        // This example specifies whether the sheet gridlines must be displayed or not.
        
        // How to set whether sheet gridlines should be displayed or not.
        
        // Set a boolean value representing whether to display gridlines or not.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A2").SetValue("The sheet settings make it display no gridlines");
        worksheet.SetDisplayGridlines(false);
        
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
