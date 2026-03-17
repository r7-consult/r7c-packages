/**
 * @file SetAuthorName_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCommentReply.SetAuthorName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the author's name for a comment reply.
 * It adds a comment to a cell, adds a reply to it, sets a new author name for the reply, and then displays the updated author name.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить имя автора для ответа на комментарий.
 * Он добавляет комментарий к ячейке, добавляет к нему ответ, устанавливает новое имя автора для ответа, а затем отображает обновленное имя автора.
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
        // This example sets the comment reply author's name.
        
        // How to add author's name to the reply.
        
        // Add a reply to the comment and set author name, then show author name in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.");
        comment.AddReply("Reply 1", "John Smith", "uid-1");
        let reply = comment.GetReply();
        reply.SetAuthorName("Mark Potato");
        worksheet.GetRange("A3").SetValue("Comment's reply author: ");
        worksheet.GetRange("B3").SetValue(reply.GetAuthorName());
        
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
