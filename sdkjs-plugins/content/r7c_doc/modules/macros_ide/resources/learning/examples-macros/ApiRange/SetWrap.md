/**
 * @file SetWrap_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetWrap
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to specify whether the words in the cell must be wrapped to fit the cell size.
 * It sets a value in cell A1, enables text wrapping for it, and then displays whether the text is wrapped.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как указать, должны ли слова в ячейке переноситься по размеру ячейки.
 * Он устанавливает значение в ячейке A1, включает перенос текста для нее, а затем отображает, переносится ли текст.
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
        // This example specifies whether the words in the cell must be wrapped to fit the cell size or not.
        
        // How to wrapp a text in the cell.
        
        // Get a range and make its content wrapped.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1");
        range.SetValue("This is the text wrapped to fit the cell size.");
        range.SetWrap(true);
        worksheet.GetRange("A3").SetValue("The text in the cell A1 is wrapped: " + range.GetWrapText());
        
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
