/**
 * @file GetText_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiComment.GetText
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the text of a comment.
 * It adds a comment to a cell and then displays the comment's text in another cell.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить текст комментария.
 * Он добавляет комментарий к ячейке, а затем отображает текст комментария в другой ячейке.
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
        // This example shows how to get the comment text.
        
        // How to get a comment raw text.
        
        // Add a comment text to a range of the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        range.AddComment("This is just a number.");
        worksheet.GetRange("A3").SetValue("Comment: ");
        worksheet.GetRange("B3").SetValue(range.GetComment().GetText());
        
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
