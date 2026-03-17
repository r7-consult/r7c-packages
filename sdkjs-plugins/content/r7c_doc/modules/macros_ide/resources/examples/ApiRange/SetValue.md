/**
 * @file SetValue_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetValue
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set a value to cells.
 * It sets numerical values in cells B1 and B2, a text label in A3, and a formula in B3 to calculate the product of B1 and B2.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить значение в ячейках.
 * Он устанавливает числовые значения в ячейках B1 и B2, текстовую метку в A3 и формулу в B3 для вычисления произведения B1 и B2.
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
        // This example sets a value to cells.
        
        // How to add underline to the cell value.
        
        // Get a range and add underline its text.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B1").SetValue("2");
        worksheet.GetRange("B2").SetValue("2");
        worksheet.GetRange("A3").SetValue("2x2=");
        worksheet.GetRange("B3").SetValue("=B1*B2");
        
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
