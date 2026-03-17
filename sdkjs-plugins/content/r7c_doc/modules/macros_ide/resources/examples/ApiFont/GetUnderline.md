/**
 * @file GetUnderline_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiFont.GetUnderline
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the type of underline applied to a specified font.
 * It sets a sample text in cell B1, applies a single underline to a part of it, and then displays the underline type.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить тип подчеркивания, примененного к указанному шрифту.
 * Он устанавливает образец текста в ячейке B1, применяет к части его одинарное подчеркивание, а затем отображает тип подчеркивания.
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
        // This example shows how to get the type of underline applied to the specified font.
        
        // How to determine whether a font is underlined or not.
        
        // Get a boolean value that represents whether a font has an underline property or not and show the value in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1");
        range.SetValue("This is just a sample text.");
        let characters = range.GetCharacters(9, 4);
        let font = characters.GetFont();
        font.SetUnderline("xlUnderlineStyleSingle");
        let underlineType = font.GetUnderline();
        worksheet.GetRange("B3").SetValue("Underline property: " + underlineType);
        
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
