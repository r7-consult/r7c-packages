/**
 * @file SetUserId_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiComment.SetUserId
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the user ID for a comment's author.
 * It adds a comment to a cell, sets a new user ID, and then displays the updated user ID.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить идентификатор пользователя для автора комментария.
 * Он добавляет комментарий к ячейке, устанавливает новый идентификатор пользователя, а затем отображает обновленный идентификатор пользователя.
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
        // This example sets the user ID to the comment author.
        
        // How to change a comment author ID.
        
        // Replace a comment author ID to a new one.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.", "John Smith");
        worksheet.GetRange("A3").SetValue("Comment's user Id: ");
        comment.SetUserId("uid-2");
        worksheet.GetRange("B3").SetValue(comment.GetUserId());
        
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
