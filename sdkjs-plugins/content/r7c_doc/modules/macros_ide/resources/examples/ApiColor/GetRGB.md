/**
 * @file GetRGB_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiColor.GetRGB
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the RGB value of a color object.
 * It creates a color from an RGB value, applies it to a cell's font, and then displays the RGB value of the color object.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить значение RGB объекта цвета.
 * Он создает цвет из значения RGB, применяет его к шрифту ячейки, а затем отображает значение RGB объекта цвета.
 *
 * @returns {void}
 *
 * @see https://r7-consult.ru/
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
        // This example gets an RGB format of a color and inserts it into the table.
        
        // How to get a RGB color format.
        
        // Convert a color to the RGB values.
        
        let worksheet = Api.GetActiveSheet();
        let color = Api.CreateColorFromRGB(255, 111, 61);
        worksheet.GetRange("A2").SetValue("Text with color");
        worksheet.GetRange("A2").SetFontColor(color);
        let rgbColor = color.GetRGB();
        worksheet.GetRange("A4").SetValue("Cell color in RGB format: " + rgbColor);
        
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
