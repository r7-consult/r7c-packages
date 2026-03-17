/**
 * @file GetWrapText_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetWrapText
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the wrap text property of a range.
 * It sets a value in cell A1, sets its wrap text property to true, and then displays whether the text is wrapped.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить свойство переноса текста диапазона.
 * Он устанавливает значение в ячейке A1, устанавливает для него свойство переноса текста в значение true, а затем отображает, переносится ли текст.
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
        // This example shows how to get the information about the wrapping cell style.
        
        // How to get a cell value wrapping type.
        
        // Get a wrapping type of a cell from its range and show it.
        
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
