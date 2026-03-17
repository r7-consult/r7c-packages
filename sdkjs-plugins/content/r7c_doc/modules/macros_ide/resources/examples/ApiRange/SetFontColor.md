/**
 * @file SetFontColor_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetFontColor
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the text color to a cell range.
 * It sets the font color of cell A2 to an orange color and then sets a value in it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить цвет текста для диапазона ячеек.
 * Он устанавливает цвет шрифта ячейки A2 в оранжевый цвет, а затем устанавливает в ней значение.
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
        // This example sets the text color to the cell range.
        
        // How to color a cell text.
        
        // Get a range and apply an RGB color to its text color.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A2").SetFontColor(Api.CreateColorFromRGB(255, 111, 61));
        worksheet.GetRange("A2").SetValue("This is the text with a color set to it");
        worksheet.GetRange("A4").SetValue("This is the text with a default color");
        
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
