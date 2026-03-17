/**
 * @file SetFontSize_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetFontSize
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the font size to the characters of a cell range.
 * It sets a value in cell A2 and then sets the font size of the range A1:D5 to 20.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить размер шрифта для символов диапазона ячеек.
 * Он устанавливает значение в ячейке A2, а затем устанавливает размер шрифта диапазона A1:D5 на 20.
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
        // This example sets the font size to the characters of the cell range.
        
        // How to resize a cell font size.
        
        // Get a range and set its font size.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A2").SetValue("2");
        let range = worksheet.GetRange("A1:D5");
        range.SetFontSize(20);
        
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
