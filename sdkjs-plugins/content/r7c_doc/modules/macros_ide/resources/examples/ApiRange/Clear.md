/**
 * @file Clear_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.Clear
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to clear the content of a range.
 * It sets a value in range A1:B1 and then clears everything from it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как очистить содержимое диапазона.
 * Он устанавливает значение в диапазоне A1:B1, а затем очищает из него все.
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
        // This example clears the range.
        
        // How to clear a content of a range.
        
        // Get a range and remove everything from it.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1:B1");
        range.SetValue("1");
        range.Clear();
        worksheet.GetRange("A2").SetValue("The range A1:B1 was just cleared.");
        
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
