/**
 * @file SetColor_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiFont.SetColor
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the font color property of a specified font.
 * It sets a sample text in cell B1 and then sets the color of a part of it to an RGB color.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить свойство цвета указанного шрифта.
 * Он устанавливает образец текста в ячейке B1, а затем устанавливает цвет его части в цвет RGB.
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
        // This example sets the font color property to the specified font.
        
        // How to change a text color.
        
        // Get a font object of characters and color it specifying a color in RGB format.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(9, 4);
        let font = characters.GetFont();
        let color = Api.CreateColorFromRGB(255, 111, 61);
        font.SetColor(color);
        
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
