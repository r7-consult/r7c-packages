/**
 * @file Delete_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCharacters.Delete
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to delete a range of characters from a cell in a spreadsheet.
 * It sets a sample text in cell B1, then deletes 4 characters starting from the 9th position.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как удалить диапазон символов из ячейки в электронной таблице.
 * Он устанавливает образец текста в ячейке B1, а затем удаляет 4 символа, начиная с 9-й позиции.
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
        // This example deletes the ApiCharacters object.
        
        // How to delete characters from an array.
        
        // Remove all characters.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(9, 4);
        characters.Delete();
        
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
