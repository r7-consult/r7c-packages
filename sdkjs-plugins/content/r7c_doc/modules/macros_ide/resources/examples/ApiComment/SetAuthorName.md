/**
 * @file SetAuthorName_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiComment.SetAuthorName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the author's name for a comment.
 * It adds a comment to a cell, sets a new author name, and then displays the updated author name.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить имя автора для комментария.
 * Он добавляет комментарий к ячейке, устанавливает новое имя автора, а затем отображает обновленное имя автора.
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
        // This example sets the comment author's name.
        
        // How to add author's name to the comment.
        
        // Add a comment and author name to it, then show author name in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.", "John Smith");
        worksheet.GetRange("A3").SetValue("Comment's author: ");
        comment.SetAuthorName("Mark Potato");
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
