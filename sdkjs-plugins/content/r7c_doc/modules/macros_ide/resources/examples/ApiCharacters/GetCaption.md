/**
 * @file GetCaption_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCharacters.GetCaption
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the caption of a range of characters from a cell in a spreadsheet.
 * It sets a sample text in cell B1, then gets the caption of 4 characters starting from the 23rd position and displays it in cell B3.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить заголовок диапазона символов из ячейки в электронной таблице.
 * Он устанавливает образец текста в ячейке B1, затем получает заголовок 4 символов, начиная с 23-й позиции, и отображает его в ячейке B3.
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
        // This example shows how to get a string value that represents the text of the specified range of characters.
        
        // Get a value that represents the label text for the pivot field.
        
        // How to get and display caption of the text.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(23, 4);
        let caption = characters.GetCaption();
        worksheet.GetRange("B3").SetValue("Caption: " + caption);
        
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
