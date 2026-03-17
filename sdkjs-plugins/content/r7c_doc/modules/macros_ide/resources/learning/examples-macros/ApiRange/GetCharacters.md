/**
 * @file GetCharacters_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetCharacters
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the ApiCharacters object that represents a range of characters within the object text.
 * It sets a value in cell B1, gets a range of characters from it, and then makes their font bold.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект ApiCharacters, представляющий диапазон символов в тексте объекта.
 * Он устанавливает значение в ячейке B1, получает из него диапазон символов, а затем делает их шрифт жирным.
 *
 * @returns {void}
 *
 * @see https://r7-consult.com/
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
        // This example shows how to get the ApiCharacters object that represents a range of characters within the object text.
        
        // How to get range characters.
        
        // Get the range characters, get their font object and set it to bold.
        
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
