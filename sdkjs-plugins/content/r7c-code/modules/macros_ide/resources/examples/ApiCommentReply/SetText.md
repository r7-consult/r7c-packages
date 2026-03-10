/**
 * @file SetText_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCommentReply.SetText
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the text of a comment reply.
 * It adds a comment to a cell, adds a reply to it, sets a new text for the reply, and then displays the updated text.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить текст ответа на комментарий.
 * Он добавляет комментарий к ячейке, добавляет к нему ответ, устанавливает новый текст для ответа, а затем отображает обновленный текст.
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
        // This example sets the comment reply text.
        
        // How to change a reply text.
        
        // Replace a reply text with a new text.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.");
        comment.AddReply("Reply 1", "John Smith", "uid-1");
        let reply = comment.GetReply();
        reply.SetText("New reply text.");
        worksheet.GetRange("A3").SetValue("Comment's reply text: ");
        worksheet.GetRange("B3").SetValue(reply.GetText());
        
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
