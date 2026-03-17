/**
 * @file Select_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.Select
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to select a range.
 * It sets a value in range A1:C1, selects it, and then sets the value of the current selection to "selected".
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как выбрать диапазон.
 * Он устанавливает значение в диапазоне A1:C1, выбирает его, а затем устанавливает значение текущего выделения в «selected».
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
        // This example selects the current range.
        
        // How to select a range.
        
        // Select a range and get a selection from the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1:C1");
        range.SetValue("1");
        range.Select();
        Api.GetSelection().SetValue("selected");
        
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
