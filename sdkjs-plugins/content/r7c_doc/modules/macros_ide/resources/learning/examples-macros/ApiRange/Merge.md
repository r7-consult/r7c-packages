/**
 * @file Merge_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.Merge
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to merge a selected cell range into a single cell or a cell row.
 * It merges range A3:E8 into a single cell and range A9:E14 into a single row.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как объединить выбранный диапазон ячеек в одну ячейку или строку ячеек.
 * Он объединяет диапазон A3:E8 в одну ячейку и диапазон A9:E14 в одну строку.
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
        // This example merges the selected cell range into a single cell or a cell row.
        
        // How to get a range using cell addresses and merge them into one.
        
        // Get a range, merge its cells into one cell.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A3:E8").Merge(true);
        worksheet.GetRange("A9:E14").Merge(false);
        
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
