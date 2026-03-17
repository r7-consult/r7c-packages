/**
 * @file SetCaption_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCharacters.SetCaption
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the caption for a specified range of characters.
 * It sets a sample text in cell B1, then sets the caption "string" for 4 characters starting from the 23rd position.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить заголовок для указанного диапазона символов.
 * Он устанавливает образец текста в ячейке B1, а затем устанавливает заголовок "string" для 4 символов, начиная с 23-й позиции.
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
        // This example sets a string value that represents the text of the specified range of characters.
        
        // How to add a label for the specified characters.
        
        // Set a caption for the characters collection.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(23, 4);
        characters.SetCaption("string");
        
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
