/**
 * @file SetText_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCharacters.SetText
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the text for a specified range of characters.
 * It sets a sample text in cell B1, then sets the text for 4 characters starting from the 23rd position to "string".
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить текст для указанного диапазона символов.
 * Он устанавливает образец текста в ячейке B1, а затем устанавливает текст для 4 символов, начиная с 23-й позиции, в "string".
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
        // This example sets the text for the specified characters.
        
        // Update characters collection by setting a new text.
        
        // Set text for the characters of the range.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(23, 4);
        characters.SetText("string");
        
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
