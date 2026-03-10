/**
 * @file SetOrientation_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetOrientation
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the orientation of a cell range.
 * It sets values in a range and then sets its orientation to "xlUpward".
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить ориентацию диапазона ячеек.
 * Он устанавливает значения в диапазоне, а затем устанавливает его ориентацию на «xlUpward».
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
        // This example sets an angle to the cell range.
        
        // How to set an orientation of cells.
        
        // Get a range and specify its orientation.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        let range = worksheet.GetRange("A1:B1");
        range.SetOrientation("xlUpward");
        
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
