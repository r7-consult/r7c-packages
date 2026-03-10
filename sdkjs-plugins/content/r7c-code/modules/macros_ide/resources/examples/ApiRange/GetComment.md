/**
 * @file GetComment_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetComment
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the ApiComment object of a range.
 * It sets a value in cell A1, adds a comment to it, and then displays the comment's text.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект ApiComment диапазона.
 * Он устанавливает значение в ячейке A1, добавляет к нему комментарий, а затем отображает текст комментария.
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
        // This example shows how to get the ApiComment object of the range.
        
        // How to get a range comment.
        
        // Get a range, get its comment and show its text in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1");
        range.SetValue("1");
        range.AddComment("This is just a number.");
        worksheet.GetRange("A3").SetValue("Comment: " + range.GetComment().GetText());
        
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
