/**
 * @file GetCells_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetCells
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the ApiRange that represents all the cells on the worksheet.
 * It retrieves all cells from the worksheet and then sets their fill color.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить ApiRange, представляющий все ячейки на листе.
 * Он извлекает все ячейки из листа, а затем устанавливает их цвет заливки.
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
        // This example shows how to get the ApiRange that represents all the cells on the worksheet.
        
        // How to get all cells.
        
        // Get all cells from the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let cells = worksheet.GetCells();
        cells.SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        
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
