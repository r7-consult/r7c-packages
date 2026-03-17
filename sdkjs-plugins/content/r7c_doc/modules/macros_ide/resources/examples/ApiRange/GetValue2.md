/**
 * @file GetValue2_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetValue2
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the value without format of a specified range.
 * It sets a formatted value in cell A1, gets its raw value, and then displays it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить значение без формата указанного диапазона.
 * Он устанавливает отформатированное значение в ячейке A1, получает его необработанное значение, а затем отображает его.
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
        // This example shows how to get the value without format of the specified range.
        
        // How to get a cell raw value.
        
        // Get a range, get its raw value without format and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let format = Api.Format("123456", "$#,##0");
        let range = worksheet.GetRange("A1");
        range.SetValue(format);
        let value2 = range.GetValue2();
        worksheet.GetRange("A3").SetValue("Value of the cell A1 without format: " + value2);
        
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
