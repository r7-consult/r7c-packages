/**
 * @file GetItalic_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiFont.GetItalic
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the italic property of a specified font.
 * It sets a sample text in cell B1, makes a part of it italic, and then displays whether it is italic.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить свойство курсива указанного шрифта.
 * Он устанавливает образец текста в ячейке B1, делает часть его курсивом, а затем отображает, является ли он курсивом.
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
        // This example shows how to get the italic property of the specified font.
        
        // How to know whether a font style of characters is italic.
        
        // Get a boolean value that represents whether a font is italic or not and show the value in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(9, 4);
        let font = characters.GetFont();
        font.SetItalic(true);
        let isItalic = font.GetItalic();
        worksheet.GetRange("B3").SetValue("Italic property: " + isItalic);
        
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
