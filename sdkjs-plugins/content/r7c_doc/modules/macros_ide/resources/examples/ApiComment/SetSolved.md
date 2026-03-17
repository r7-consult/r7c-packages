/**
 * @file SetSolved_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiComment.SetSolved
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to mark a comment as solved.
 * It adds a comment to a cell, marks it as solved, and then displays its solved status.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как пометить комментарий как решенный.
 * Он добавляет комментарий к ячейке, помечает его как решенный, а затем отображает его статус решения.
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
        // This example marks a comment as solved.
        
        // How to resolve a comment.
        
        // Resolve a comment, then show its status in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        let comment = range.AddComment("This is just a number.", "John Smith");
        worksheet.GetRange("A3").SetValue("Comment is solved: ");
        comment.SetSolved(true);
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
