/**
 * @file SetFontName_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetFontName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the specified font family as the font name for a cell range.
 * It sets a value in cell A2 and then sets the font name of the range A1:D5 to "Arial".
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить указанное семейство шрифтов в качестве имени шрифта для диапазона ячеек.
 * Он устанавливает значение в ячейке A2, а затем устанавливает имя шрифта диапазона A1:D5 на «Arial».
 *
 * @returns {void}
 *
 * @see https://r7-consult.com/
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
        // This example sets the specified font family as the font name for the cell range.
        
        // How to change a cell font family.
        
        // Get a range and set its font family using its name.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A2").SetValue("2");
        let range = worksheet.GetRange("A1:D5");
        range.SetFontName("Arial");
        
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
