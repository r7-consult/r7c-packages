/**
 * @file SetBold_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiFont.SetBold
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the bold property of a specified font.
 * It sets a sample text in cell B1 and then makes a part of it bold.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить свойство жирности указанного шрифта.
 * Он устанавливает образец текста в ячейке B1, а затем делает часть его жирной.
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
        // This example sets the bold property to the specified font.
        
        // How to make a text bold.
        
        // Get a font object of characters and make it bold.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(9, 4);
        let font = characters.GetFont();
        font.SetBold(true);
        
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
