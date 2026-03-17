/**
 * @file Copy_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.Copy
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to copy a range to a specified destination.
 * It sets a value in cell A1 and then copies it to cell A3.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как скопировать диапазон в указанное место назначения.
 * Он устанавливает значение в ячейке A1, а затем копирует его в ячейку A3.
 *
 * @returns {void}
 *
 * @see https://r7-consult.ru/
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
        // This example copies a range to the specified range.
        
        // How to create identical range.
        
        // Get a range and create a copy of it.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1");
        range.SetValue("This is a sample text which is copied to the range A3.");
        range.Copy(worksheet.GetRange("A3"));
        
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
