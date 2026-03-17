/**
 * @file Move_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.Move
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to move the sheet to another location in the workbook.
 * It adds a new sheet and then moves it before the first sheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как переместить лист в другое место в рабочей книге.
 * Он добавляет новый лист, а затем перемещает его перед первым листом.
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
        // This example moves the sheet to another location in the workbook.
        
        // How to change an order of the sheet.
        
        // Move a sheet.
        
        let sheet1 = Api.GetActiveSheet();
        Api.AddSheet("Sheet2");
        let sheet2 = Api.GetActiveSheet();
        sheet2.Move(sheet1);
        
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
