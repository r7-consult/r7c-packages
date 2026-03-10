/**
 * @file Cut_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.Cut
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to cut a range to the clipboard.
 * It sets a value in cell A1 and then cuts it to cell A3.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вырезать диапазон в буфер обмена.
 * Он устанавливает значение в ячейке A1, а затем вырезает его в ячейку A3.
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
        // This example cuts a range to the clipboard.
        
        // How to cut a range.
        
        // Get a range, set some value for it and cut it to the clipboard.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1");
        range.SetValue("This is a sample text which is move to the range A3.");
        range.Cut(worksheet.GetRange("A3"));
        
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
