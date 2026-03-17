/**
 * @file GetClassType_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiComment.GetClassType
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the class type of a comment.
 * It adds a comment to a cell and then displays the class type of the comment object.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить тип класса комментария.
 * Он добавляет комментарий к ячейке, а затем отображает тип класса объекта комментария.
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
        
        // How to get a comment class type.
        
        // Get an comment class type to show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let range = worksheet.GetRange("A1");
        range.AddComment("This is just a number.");
        let comment = range.GetComment();
        let type = comment.GetClassType();
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
