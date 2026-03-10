/**
 * @file GetAuthorName_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiComment.GetAuthorName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the author's name from a comment.
 * It adds a comment to a cell and then displays the author's name in another cell.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить имя автора из комментария.
 * Он добавляет комментарий к ячейке, а затем отображает имя автора в другой ячейке.
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
        // This example shows how to get the comment author's name.
        
        // How to remove a comment from a range.
        
        // Get a range, add a comment to it and then remove it.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.");
        worksheet.GetRange("A3").SetValue("Comment's author: ");
        worksheet.GetRange("B3").SetValue(comment.GetAuthorName());
        
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
