/**
 * @file UnMerge_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.UnMerge
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to split a selected merged cell range into single cells.
 * It merges range A3:E8 and then unmerges range A5:E5.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как разделить выбранный объединенный диапазон ячеек на отдельные ячейки.
 * Он объединяет диапазон A3:E8, а затем разъединяет диапазон A5:E5.
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
        // This example splits the selected merged cell range into the single cells.
        
        // How to unmerge a range of cells.
        
        // Get a range and split its merged cells.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A3:E8").Merge(true);
        worksheet.GetRange("A5:E5").UnMerge();
        
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
