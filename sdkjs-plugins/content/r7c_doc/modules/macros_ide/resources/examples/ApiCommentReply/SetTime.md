/**
 * @file SetTime_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCommentReply.SetTime
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the timestamp of a comment reply's creation.
 * It adds a comment to a cell, adds a reply to it, sets its timestamp to the current time, and then displays the updated timestamp.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить временную метку создания ответа на комментарий.
 * Он добавляет комментарий к ячейке, добавляет к нему ответ, устанавливает его временную метку на текущее время, а затем отображает обновленную временную метку.
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
        // This example sets the timestamp of the comment reply creation in the current time zone format.
        
        // How to change a time when a reply was created.
        
        // Add a reply then update its creation time and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.");
        comment.AddReply("Reply 1", "John Smith", "uid-1");
        let reply = comment.GetReply();
        reply.SetTime(Date.now());
        worksheet.GetRange("A3").SetValue("Comment's reply timestamp: ");
        worksheet.GetRange("B3").SetValue(reply.GetTime());
        
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
