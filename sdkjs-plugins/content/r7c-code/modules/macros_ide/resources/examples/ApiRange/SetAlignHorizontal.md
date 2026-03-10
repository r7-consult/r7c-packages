/**
 * @file SetAlignHorizontal_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetAlignHorizontal
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the horizontal alignment of the text in a cell range.
 * It sets a value in cell A2 and then sets the horizontal alignment of the range A1:D5 to "center".
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить горизонтальное выравнивание текста в диапазоне ячеек.
 * Он устанавливает значение в ячейке A2, а затем устанавливает горизонтальное выравнивание диапазона A1:D5 по центру.
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
        // This example sets the horizontal alignment of the text in the cell range.
        
        // How to change the horizontal alignment of the cell content.
        
        // Change the horizontal alignment of the ApiRange content to center.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A2").SetValue("2");
        let range = worksheet.GetRange("A1:D5");
        range.SetAlignHorizontal("center");
        
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
