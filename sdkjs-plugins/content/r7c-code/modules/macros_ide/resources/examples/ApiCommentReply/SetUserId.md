/**
 * @file SetUserId_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCommentReply.SetUserId
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the user ID for a comment reply's author.
 * It adds a comment to a cell, adds a reply to it, sets a new user ID for the reply, and then displays the updated user ID.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить идентификатор пользователя для автора ответа на комментарий.
 * Он добавляет комментарий к ячейке, добавляет к нему ответ, устанавливает новый идентификатор пользователя для ответа, а затем отображает обновленный идентификатор пользователя.
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
        // This example sets the user ID to the comment reply author.
        
        // How to change a reply author ID.
        
        // Replace a reply author ID to a new one.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.");
        comment.AddReply("Reply 1", "John Smith", "uid-1");
        let reply = comment.GetReply();
        reply.SetUserId("uid-2");
        worksheet.GetRange("A3").SetValue("Comment's reply user Id: ");
        worksheet.GetRange("B3").SetValue(reply.GetUserId());
        
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
