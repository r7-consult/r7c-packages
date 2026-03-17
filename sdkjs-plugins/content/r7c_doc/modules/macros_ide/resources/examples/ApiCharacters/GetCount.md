/**
 * @file GetCount_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCharacters.GetCount
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the number of characters in a specified range within a cell.
 * It sets a sample text in cell B1, then gets the number of characters in a range of 4 characters starting from the 23rd position and displays it in cell B3.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить количество символов в указанном диапазоне внутри ячейки.
 * Он устанавливает образец текста в ячейке B1, затем получает количество символов в диапазоне из 4 символов, начиная с 23-й позиции, и отображает его в ячейке B3.
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
        // This example shows how to get a value that represents a number of objects in the collection.
        
        // How to get collection objects count.
        
        // How to get array length.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(23, 4);
        let count = characters.GetCount();
        worksheet.GetRange("B3").SetValue("Number of characters: " + count);
        
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
