/**
 * @file RemoveReplies_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiComment.RemoveReplies
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to remove replies from a comment.
 * It adds a comment to a cell, adds two replies, removes the first reply, and then displays the remaining replies count.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как удалить ответы из комментария.
 * Он добавляет комментарий к ячейке, добавляет два ответа, удаляет первый ответ, а затем отображает количество оставшихся ответов.
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
        // This example removes the specified comment replies.
        
        // How to remove replies from the comment.
        
        // Add a comment and replies to it, then remove specified reply and then show replies count.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.");
        comment.AddReply("Reply 1", "John Smith", "uid-1");
        comment.AddReply("Reply 2", "John Smith", "uid-1");
        comment.RemoveReplies(0, 1, false);
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
