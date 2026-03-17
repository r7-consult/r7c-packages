/**
 * @file GetValue_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetValue
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the value of a specified range.
 * It sets a value in cell A1, gets its value, and then displays it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить значение указанного диапазона.
 * Он устанавливает значение в ячейке A1, получает его значение, а затем отображает его.
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
        // This example shows how to get a value of the specified range.
        
        // How to get a cell value.
        
        // Get a range, get its value and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        let value = worksheet.GetRange("A1").GetValue();
        worksheet.GetRange("A3").SetValue("Value of the cell A1: ");
        worksheet.GetRange("B3").SetValue(value);
        
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
