/**
 * @file FormatAsTable_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.FormatAsTable
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to format the selected range of cells from the sheet as a table.
 * It formats the range A1:E10 as a table.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как отформатировать выбранный диапазон ячеек листа как таблицу.
 * Он форматирует диапазон A1:E10 как таблицу.
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
        // This example formats the selected range of cells from the sheet as a table.
        
        // How to format a range as a table.
        
        // Select a range and format it as a table.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.FormatAsTable("A1:E10");
        
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
