/**
 * @file Paste_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.Paste
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to paste a range to a specified destination.
 * It sets values in range B4:D4 and then pastes them to range A1:C1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вставить диапазон в указанное место назначения.
 * Он устанавливает значения в диапазоне B4:D4, а затем вставляет их в диапазон A1:C1.
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
        // This example pastes the Range object to the specified range.
        
        // How to get a range and paste it into another one.
        
        // Create a range and add it to another one.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B4").SetValue("1");
        worksheet.GetRange("C4").SetValue("2");
        worksheet.GetRange("D4").SetValue("3");
        let rangeFrom = worksheet.GetRange("B4:D4");
        let range = worksheet.GetRange("A1:C1");
        range.Paste(rangeFrom);
        
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
