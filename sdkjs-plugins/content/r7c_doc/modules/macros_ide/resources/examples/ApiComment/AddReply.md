/**
 * @file AddReply_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiComment.AddReply
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add a reply to a comment.
 * It adds a comment to a cell, then adds a reply to that comment and displays the reply text in another cell.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавить ответ на комментарий.
 * Он добавляет комментарий к ячейке, затем добавляет ответ на этот комментарий и отображает текст ответа в другой ячейке.
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
        // This example adds a reply to a comment.
        
        // How to reply to a comment.
        
        // Add a commnet reply indicating an author and id.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.");
        comment.AddReply("Reply 1", "John Smith", "uid-1");
        let reply = comment.GetReply();
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
