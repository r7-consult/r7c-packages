/**
 * @file SetTimeUTC_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiComment.SetTimeUTC
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the UTC timestamp of a comment's creation.
 * It adds a comment to a cell, sets its UTC timestamp to the current time, and then displays the updated UTC timestamp.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить временную метку UTC создания комментария.
 * Он добавляет комментарий к ячейке, устанавливает его временную метку UTC на текущее время, а затем отображает обновленную временную метку UTC.
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
        // This example sets the timestamp of the comment creation in UTC format.
        
        // How to change a timestamp in UTC when a comment was created.
        
        // Add a comment then update its creation time in UTC format and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.", "John Smith");
        worksheet.GetRange("A3").SetValue("Timestamp UTC: ");
        comment.SetTimeUTC(Date.now());
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
