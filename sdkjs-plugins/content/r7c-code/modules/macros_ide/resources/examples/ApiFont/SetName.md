/**
 * @file SetName_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiFont.SetName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the font name property of a specified font.
 * It sets a sample text in cell B1, sets the font name of a part of it, and then displays the font name.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить свойство имени указанного шрифта.
 * Он устанавливает образец текста в ячейке B1, устанавливает имя шрифта его части, а затем отображает имя шрифта.
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
        // This example sets the font name property to the specified font.
        
        // How to change name of a font.
        
        // Get a font object of characters and set a new name for it.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(9, 4);
        let font = characters.GetFont();
        font.SetName("Font 1");
        let fontName = font.GetName();
        worksheet.GetRange("B3").SetValue("Font name: " + fontName);
        
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
