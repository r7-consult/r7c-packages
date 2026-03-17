/**
 * @file SetRowHeight_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.SetRowHeight
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the height of the specified row measured in points.
 * It sets the height of row 0 to 30 points.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить высоту указанной строки, измеряемую в пунктах.
 * Он устанавливает высоту строки 0 на 30 пунктов.
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
        // This example sets the height of the specified row measured in points.
        
        // How to resize the height of the row.
        
        // Set a row height.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetRowHeight(0, 30);
        
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
