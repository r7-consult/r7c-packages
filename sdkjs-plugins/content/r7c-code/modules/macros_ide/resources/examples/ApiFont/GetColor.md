/**
 * @file GetColor_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiFont.GetColor
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the color property of a specified font.
 * It sets a sample text in cell B1, sets the color of a part of it, and then applies that color to another part of the text.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить свойство цвета указанного шрифта.
 * Он устанавливает образец текста в ячейке B1, устанавливает цвет его части, а затем применяет этот цвет к другой части текста.
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
        // This example shows how to get the font color property of the specified font.
        
        // How to know a font color of the characters.
        
        // Get a color value represented in RGB format and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(9, 4);
        let font = characters.GetFont();
        let color = Api.CreateColorFromRGB(255, 111, 61);
        font.SetColor(color);
        color = font.GetColor();
        characters = range.GetCharacters(16, 6);
        font = characters.GetFont();
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
