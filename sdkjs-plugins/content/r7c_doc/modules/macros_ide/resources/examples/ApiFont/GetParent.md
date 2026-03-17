/**
 * @file GetParent_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiFont.GetParent
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the parent ApiCharacters object of a specified font.
 * It sets a sample text in cell B1, gets the font of a part of it, and then sets the text of the parent characters object.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить родительский объект ApiCharacters указанного шрифта.
 * Он устанавливает образец текста в ячейке B1, получает шрифт его части, а затем устанавливает текст родительского объекта символов.
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
        // This example shows how to get the parent ApiCharacters object of the specified font.
        
        // How to determine a font object's parent.
        
        // Get a parent of a font and add text to it.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(23, 4);
        let font = characters.GetFont();
        let parent = font.GetParent();
        parent.SetText("string");
        
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
