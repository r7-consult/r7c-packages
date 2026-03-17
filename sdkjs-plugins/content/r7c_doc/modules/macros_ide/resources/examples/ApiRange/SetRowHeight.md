/**
 * @file SetRowHeight_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetRowHeight
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the row height value.
 * It sets the height of row 1 to 32.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить значение высоты строки.
 * Он устанавливает высоту строки 1 на 32.
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
        // This example sets the row height value.
        
        // How to set a row height of cells.
        
        // Get a range and specify its row height.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetRowHeight(32);
        
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
