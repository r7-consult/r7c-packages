/**
 * @file GetTimeUTC_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiComment.GetTimeUTC
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the UTC timestamp of a comment's creation.
 * It adds a comment to a cell and then displays the UTC timestamp in another cell.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить временную метку UTC создания комментария.
 * Он добавляет комментарий к ячейке, а затем отображает временную метку UTC в другой ячейке.
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
        // This example shows how to get the timestamp of the comment creation in UTC format.
        
        // How to get a time in UTC when a comment was created.
        
        // Add a comment creation timestamp UTC to a range of the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.");
        worksheet.GetRange("A3").SetValue("Timestamp UTC: ");
        worksheet.GetRange("B3").SetValue(comment.GetTimeUTC());
        
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
