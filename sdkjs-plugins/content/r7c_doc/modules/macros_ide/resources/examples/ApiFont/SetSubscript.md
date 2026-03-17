/**
 * @file SetSubscript_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiFont.SetSubscript
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the subscript property of a specified font.
 * It sets a sample text in cell B1 and then makes a part of it subscript.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить свойство подстрочного индекса указанного шрифта.
 * Он устанавливает образец текста в ячейке B1, а затем делает часть его подстрочным индексом.
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
        // This example sets the subscript property to the specified font.
        
        // How to change a regular text to a subscript.
        
        // Get a font object of characters and make it subscript.
        
        const worksheet = Api.GetActiveSheet();
        const range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        const characters = range.GetCharacters(9, 4);
        const font = characters.GetFont();
        font.SetSubscript(true);
        
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
