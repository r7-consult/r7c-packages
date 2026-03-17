/**
 * @file Delete_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.Delete
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to delete a worksheet.
 * It adds a new sheet, then deletes it, and displays a confirmation message.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как удалить лист.
 * Он добавляет новый лист, затем удаляет его и отображает сообщение с подтверждением.
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
        // This example deletes the worksheet.
        
        // How to delete sheets.
        
        // Remove a worksheet.
        
        Api.AddSheet("New sheet");
        let sheet = Api.GetActiveSheet();
        sheet.Delete();
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A3").SetValue("This method just deleted the second sheet from this spreadsheet.");
        
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
