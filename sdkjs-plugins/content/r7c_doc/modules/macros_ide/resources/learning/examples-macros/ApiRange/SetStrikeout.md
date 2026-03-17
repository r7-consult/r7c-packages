/**
 * @file SetStrikeout_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetStrikeout
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to specify that the contents of the cell are displayed with a single horizontal line through the center of the contents.
 * It sets a value in cell A2 and then applies strikethrough to it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как указать, что содержимое ячейки отображается с одной горизонтальной линией через центр содержимого.
 * Он устанавливает значение в ячейке A2, а затем применяет к нему зачеркивание.
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
        // This example specifies that the contents of the cell is displayed with a single horizontal line through the center of the contents.
        
        // How to add strikeout to the cell value.
        
        // Get a range and add strikeout to its text.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A2").SetValue("Struckout text");
        worksheet.GetRange("A2").SetStrikeout(true);
        worksheet.GetRange("A3").SetValue("Normal text");
        
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
