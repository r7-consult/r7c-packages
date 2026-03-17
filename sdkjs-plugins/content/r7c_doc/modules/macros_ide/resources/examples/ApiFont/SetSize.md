/**
 * @file SetSize_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiFont.SetSize
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the font size property of a specified font.
 * It sets a sample text in cell B1 and then sets the font size of a part of it to 18.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить свойство размера указанного шрифта.
 * Он устанавливает образец текста в ячейке B1, а затем устанавливает размер шрифта его части на 18.
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
        // This example sets the font size property to the specified font.
        
        // How to change the font size.
        
        // Get a font object of characters and resize it.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(9, 4);
        let font = characters.GetFont();
        font.SetSize(18);
        
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
