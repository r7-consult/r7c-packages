/**
 * @file IsSolved_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiComment.IsSolved
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to check if a comment is solved.
 * It adds a comment to a cell and then displays whether the comment is solved in another cell.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как проверить, решен ли комментарий.
 * Он добавляет комментарий к ячейке, а затем отображает, решен ли комментарий, в другой ячейке.
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
        // This example checks if a comment is solved or not.
        
        // How to find out whether a comment is resolved.
        
        // Add a comment resolved status to a range of the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.");
        worksheet.GetRange("A3").SetValue("Comment is solved: ");
        worksheet.GetRange("B3").SetValue(comment.IsSolved());
        
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
