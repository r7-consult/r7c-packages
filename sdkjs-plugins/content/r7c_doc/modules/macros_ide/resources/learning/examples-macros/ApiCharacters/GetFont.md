/**
 * @file GetFont_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCharacters.GetFont
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the font of a specified range of characters and modify it.
 * It sets a sample text in cell B1, then gets the font of 4 characters starting from the 9th position and makes them bold.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить шрифт указанного диапазона символов и изменить его.
 * Он устанавливает образец текста в ячейке B1, затем получает шрифт 4 символов, начиная с 9-й позиции, и делает их жирными.
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
        // This example shows how to get the ApiFont object that represents the font of the specified characters.
        
        // How to get font style of the array of characters.
        
        // Use font of the specified characters to set their style.
        
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
