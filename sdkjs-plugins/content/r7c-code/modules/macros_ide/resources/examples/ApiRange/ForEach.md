/**
 * @file ForEach_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.ForEach
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to execute a provided function once for each cell in a range.
 * It sets values in a range and then iterates through each cell to make values other than "1" bold.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как выполнить предоставленную функцию один раз для каждой ячейки в диапазоне.
 * Он устанавливает значения в диапазоне, а затем перебирает каждую ячейку, чтобы сделать жирными значения, отличные от «1».
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
        // This example executes a provided function once for each cell.
        
        // How to iterate through each cell from a range.
        
        // For Each cycle implementation for ApiRange.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        worksheet.GetRange("C1").SetValue("3");
        let range = worksheet.GetRange("A1:C1");
        range.ForEach(function (range) {
            let value = range.GetValue();
            if (value != "1") {
                range.SetBold(true);
            }
        });
        
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
