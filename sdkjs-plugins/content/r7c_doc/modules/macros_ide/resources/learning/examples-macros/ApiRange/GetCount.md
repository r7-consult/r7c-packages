/**
 * @file GetCount_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetCount
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the number of cells in a range.
 * It sets values in a range and then displays the count of cells in that range.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить количество ячеек в диапазоне.
 * Он устанавливает значения в диапазоне, а затем отображает количество ячеек в этом диапазоне.
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
        // This example shows how to get the cells count in the range.
        
        // How to find out how many cells a range has.
        
        // Get a range, get its cells count and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        worksheet.GetRange("C1").SetValue("3");
        let count = worksheet.GetRange("A1:C1").GetCount();
        worksheet.GetRange("A4").SetValue("Count: ");
        worksheet.GetRange("B4").SetValue(count);
        
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
