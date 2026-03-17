/**
 * @file GetActiveCell_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetActiveCell
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get an object that represents an active cell.
 * It gets the active cell and then sets a value in it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект, представляющий активную ячейку.
 * Он получает активную ячейку, а затем устанавливает в ней значение.
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
        // This example shows how to get an object that represents an active cell.
        
        // How to get selected active cell.
        
        // Get an active cell and insert data to it.
        
        let worksheet = Api.GetActiveSheet();
        let activeCell = worksheet.GetActiveCell();
        activeCell.SetValue("This sample text was placed in an active cell.");
        
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
