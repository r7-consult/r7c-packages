/**
 * @file GetRepliesCount_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiComment.GetRepliesCount
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the number of replies to a comment.
 * It adds a comment to a cell, adds a reply to it, and then displays the number of replies in another cell.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить количество ответов на комментарий.
 * Он добавляет комментарий к ячейке, добавляет к нему ответ, а затем отображает количество ответов в другой ячейке.
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
        // This example shows how to get a number of the comment replies.
        
        // How to get a number of replies to the comment.
        
        // Add a comment to the range and display its replies count in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.");
        comment.AddReply("Reply 1", "John Smith", "uid-1");
        worksheet.GetRange("A3").SetValue("Comment replies count: ");
        worksheet.GetRange("B3").SetValue(comment.GetRepliesCount());
        
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
