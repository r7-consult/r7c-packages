/**
 * @file GetId_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiComment.GetId
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the ID of a comment.
 * It adds a comment to a cell and then displays the comment's ID in another cell.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить идентификатор комментария.
 * Он добавляет комментарий к ячейке, а затем отображает идентификатор комментария в другой ячейке.
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
        // This example shows how to get the comment ID.
        
        // How to get a comment ID.
        
        // Find a comment by its ID and display its ID.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        range.AddComment("This is just a number.");
        worksheet.GetRange("A3").SetValue("Comment: ");
        worksheet.GetRange("B3").SetValue(range.GetComment().GetId());
        
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
