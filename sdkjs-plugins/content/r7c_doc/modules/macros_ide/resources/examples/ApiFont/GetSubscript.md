/**
 * @file GetSubscript_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiFont.GetSubscript
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the subscript property of a specified font.
 * It sets a sample text in cell B1, makes a part of it subscript, and then displays whether it is subscript.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить свойство подстрочного индекса указанного шрифта.
 * Он устанавливает образец текста в ячейке B1, делает часть его подстрочным индексом, а затем отображает, является ли он подстрочным индексом.
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
        // This example shows how to get the subscript property of the specified font.
        
        // How to determine a font subscript property.
        
        // Get a boolean value that represents whether a font has a subscript property or not and show the value in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(9, 4);
        let font = characters.GetFont();
        font.SetSubscript(true);
        let isSubscript = font.GetSubscript();
        worksheet.GetRange("B3").SetValue("Subscript property: " + isSubscript);
        
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
