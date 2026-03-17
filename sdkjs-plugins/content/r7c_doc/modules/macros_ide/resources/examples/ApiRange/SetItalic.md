/**
 * @file SetItalic_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetItalic
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the italic property to the text characters in a cell.
 * It sets a value in cell A2 and then makes it italic.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить свойство курсива для текстовых символов в ячейке.
 * Он устанавливает значение в ячейке A2, а затем делает его курсивом.
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
        // This example sets the italic property to the text characters in the cell.
        
        // How to make a text value of cells italic.
        
        // Get a range and make specified cells font style italic.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A2").SetValue("Italicized text");
        worksheet.GetRange("A2").SetItalic(true);
        worksheet.GetRange("A3").SetValue("Normal text");
        
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
