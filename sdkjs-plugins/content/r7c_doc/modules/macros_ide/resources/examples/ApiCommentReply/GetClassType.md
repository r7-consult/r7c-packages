/**
 * @file GetClassType_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCommentReply.GetClassType
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the class type of a comment reply.
 * It adds a comment to a cell, adds a reply to it, and then displays the class type of the reply object.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить тип класса ответа на комментарий.
 * Он добавляет комментарий к ячейке, добавляет к нему ответ, а затем отображает тип класса объекта ответа.
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
        // This example gets a class type and inserts it into the table.
        
        // How to get a class type of ApiCommentReply.
        
        // Get a class type of ApiCommentReply and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.");
        comment.AddReply("Reply 1", "John Smith", "uid-1");
        let reply = comment.GetReply();
        let type = reply.GetClassType();
        worksheet.GetRange("A3").SetValue("Type: " + type);
        
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
