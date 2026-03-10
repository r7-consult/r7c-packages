/**
 * @file GetParent_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCharacters.GetParent
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the parent object of a specified range of characters and modify it.
 * It sets a sample text in cell B1, then gets the parent of 4 characters starting from the 23rd position and sets a bottom border for it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить родительский объект указанного диапазона символов и изменить его.
 * Он устанавливает образец текста в ячейке B1, затем получает родительский объект 4 символов, начиная с 23-й позиции, и устанавливает для него нижнюю границу.
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
        // This example shows how to get the parent object of the specified characters.
        
        // How to get a parent of the characters.
        
        // Find a characters parent of the selected range.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(23, 4);
        let parent = characters.GetParent();
        parent.SetBorders("Bottom", "Thick", Api.CreateColorFromRGB(255, 111, 61));
        
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
