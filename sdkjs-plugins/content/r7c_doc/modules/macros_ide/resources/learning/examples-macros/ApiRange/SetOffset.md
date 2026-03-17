/**
 * @file SetOffset_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetOffset
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the cell offset.
 * It sets a value in cell B3, then sets its offset by 2 rows and 2 columns, and then sets a new value in the new range.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить смещение ячейки.
 * Он устанавливает значение в ячейке B3, затем устанавливает его смещение на 2 строки и 2 столбца, а затем устанавливает новое значение в новом диапазоне.
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
        // This example sets the cell offset.
        
        // How to set an offset of cells.
        
        // Get a range and specify its cells offset.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B3").SetValue("Old Range");
        let range = worksheet.GetRange("B3");
        range.SetOffset(2, 2);
        range.SetValue("New Range");
        
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
