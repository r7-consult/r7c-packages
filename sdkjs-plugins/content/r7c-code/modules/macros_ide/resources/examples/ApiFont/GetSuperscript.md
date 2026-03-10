/**
 * @file GetSuperscript_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiFont.GetSuperscript
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the superscript property of a specified font.
 * It sets a sample text in cell B1, makes a part of it superscript, and then displays whether it is superscript.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить свойство надстрочного индекса указанного шрифта.
 * Он устанавливает образец текста в ячейке B1, делает часть его надстрочным индексом, а затем отображает, является ли он надстрочным индексом.
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
        // This example shows how to get the superscript property of the specified font.
        
        // How to determine a font superscript property.
        
        // Get a boolean value that represents whether a font has a superscript property or not and show the value in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(9, 4);
        let font = characters.GetFont();
        font.SetSuperscript(true);
        let isSuperscript = font.GetSuperscript();
        worksheet.GetRange("B3").SetValue("Superscript property: " + isSuperscript);
        
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
