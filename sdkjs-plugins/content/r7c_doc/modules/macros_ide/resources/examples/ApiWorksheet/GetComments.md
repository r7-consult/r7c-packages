/**
 * @file GetComments_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetComments
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get an array of ApiComment objects.
 * It adds a comment to a range, and then retrieves all comments from the worksheet to display the text of the first one.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить массив объектов ApiComment.
 * Он добавляет комментарий к диапазону, а затем извлекает все комментарии из листа для отображения текста первого из них.
 *
 * @returns {void}
 *
 * @see https://r7-consult.com/
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
        // This example shows how to get an array of ApiComment objects.
        
        // How to get all comments.
        
        // Get all comments from the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        range.AddComment("This is just a number.");
        let comments = worksheet.GetComments();
        worksheet.GetRange("A4").SetValue("Comment: " + comments[0].GetText());
        
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
