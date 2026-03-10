/**
 * @file SetBold_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetBold
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the bold property to the text characters in the current cell or cell range.
 * It sets a value in cell A2 and then makes it bold.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить свойство жирности для текстовых символов в текущей ячейке или диапазоне ячеек.
 * Он устанавливает значение в ячейке A2, а затем делает его жирным.
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
        // This example sets the bold property to the text characters in the current cell or cell range.
        
        // How to change the font style properties of a range making it bold.
        
        // Make characters of the ApiRange object bold.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A2").SetValue("Bold text");
        worksheet.GetRange("A2").SetBold(true);
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
