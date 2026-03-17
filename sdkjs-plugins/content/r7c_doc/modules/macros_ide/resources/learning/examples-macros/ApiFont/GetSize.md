/**
 * @file GetSize_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiFont.GetSize
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the font size property of a specified font.
 * It sets a sample text in cell B1, sets the font size of a part of it, and then displays the font size.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить свойство размера указанного шрифта.
 * Он устанавливает образец текста в ячейке B1, устанавливает размер шрифта его части, а затем отображает размер шрифта.
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
        // This example shows how to get the font size property of the specified font.
        
        // How to determine a font size of characters.
        
        // Get the size of a font and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(9, 4);
        let font = characters.GetFont();
        font.SetSize(18);
        let size = font.GetSize();
        worksheet.GetRange("B3").SetValue("Size property: " + size);
        
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
