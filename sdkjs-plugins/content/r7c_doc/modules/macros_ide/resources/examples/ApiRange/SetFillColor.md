/**
 * @file SetFillColor_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetFillColor
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the background color for a cell range.
 * It sets the background color of cell A2 to a peach color and then sets a value in it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить цвет фона для диапазона ячеек.
 * Он устанавливает цвет фона ячейки A2 в персиковый цвет, а затем устанавливает в ней значение.
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
        // This example sets the background color to the cell range with the previously created color object.
        
        // How to color a cell.
        
        // Get a range and apply a solid fill to its background color.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetColumnWidth(0, 50);
        worksheet.GetRange("A2").SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        worksheet.GetRange("A2").SetValue("This is the cell with a color set to its background");
        worksheet.GetRange("A4").SetValue("This is the cell with a default background color");
        
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
