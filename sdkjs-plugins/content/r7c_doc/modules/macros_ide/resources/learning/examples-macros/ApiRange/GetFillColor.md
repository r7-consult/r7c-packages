/**
 * @file GetFillColor_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetFillColor
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the background color for a cell range.
 * It sets a background color for cell A1, gets the color, and then applies it to cell A3.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить цвет фона для диапазона ячеек.
 * Он устанавливает цвет фона для ячейки A1, получает цвет, а затем применяет его к ячейке A3.
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
        // This example shows how to get the background color for the cell range.
        
        // How to find out a range background color.
        
        // Get a range get, set its background color using RGB value and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetColumnWidth(0, 60);
        let range = worksheet.GetRange("A1");
        range.SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        range.SetValue("This is the cell with a color set to its background.");
        let fillColor = range.GetFillColor();
        worksheet.GetRange("A3").SetValue("This is another cell with the same color set to its background");
        worksheet.GetRange("A3").SetFillColor(fillColor);
        
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
