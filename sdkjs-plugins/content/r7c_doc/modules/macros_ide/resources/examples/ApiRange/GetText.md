/**
 * @file GetText_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetText
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the text of a specified range.
 * It sets values in a range and then displays the text from the first cell in that range.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить текст указанного диапазона.
 * Он устанавливает значения в диапазоне, а затем отображает текст из первой ячейки в этом диапазоне.
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
        // This example shows how to get the text of the specified range.
        
        // How to get a cell raw text value.
        
        // Get a range, get its text value and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("text1");
        worksheet.GetRange("B1").SetValue("text2");
        worksheet.GetRange("C1").SetValue("text3");
        let range = worksheet.GetRange("A1:C1");
        let text = range.GetText();
        worksheet.GetRange("A3").SetValue("Text from the cell A1: " + text);
        
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
