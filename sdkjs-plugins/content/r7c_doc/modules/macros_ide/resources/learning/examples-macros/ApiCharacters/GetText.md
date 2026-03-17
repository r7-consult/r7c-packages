/**
 * @file GetText_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCharacters.GetText
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the text of a specified range of characters.
 * It sets a sample text in cell B1, then gets the text of 4 characters starting from the 23rd position and displays it in cell B3.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить текст указанного диапазона символов.
 * Он устанавливает образец текста в ячейке B1, затем получает текст 4 символов, начиная с 23-й позиции, и отображает его в ячейке B3.
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
        // This example shows how to get the text of the specified range of characters.
        
        // How to get a raw text from the characters.
        
        // Retrieve a text from the character collection.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(23, 4);
        let text = characters.GetText();
        worksheet.GetRange("B3").SetValue("Text: " + text);
        
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
