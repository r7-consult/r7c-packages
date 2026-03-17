/**
 * @file Insert_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCharacters.Insert
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to insert a string into a cell, replacing a specified range of characters.
 * It sets a sample text in cell B1, then replaces 4 characters starting from the 23rd position with the string "string".
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вставить строку в ячейку, заменяя указанный диапазон символов.
 * Он устанавливает образец текста в ячейке B1, а затем заменяет 4 символа, начиная с 23-й позиции, строкой "string".
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
        // This example inserts a string replacing the specified characters.
        
        // How to replace characters with a different string value.
        
        // Change the characters to another string value.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(23, 4);
        characters.Insert("string");
        
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
