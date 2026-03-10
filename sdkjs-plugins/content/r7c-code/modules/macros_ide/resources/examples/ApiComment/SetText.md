/**
 * @file SetText_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiComment.SetText
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the text of a comment.
 * It adds a comment to a cell and then changes its text to "New comment text".
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить текст комментария.
 * Он добавляет комментарий к ячейке, а затем изменяет его текст на "New comment text".
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
        // This example sets the comment text.
        
        // How to change a comment text.
        
        // Replace a comment text with a new text.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.");
        comment.SetText("New comment text");
        
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
