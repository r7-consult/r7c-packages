/**
 * @file SetUnderline_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetUnderline
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to specify that the contents of the current cell are displayed along with a line appearing directly below the character.
 * It sets a value in cell A2 and then applies a single underline to it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как указать, что содержимое текущей ячейки отображается вместе с линией, появляющейся непосредственно под символом.
 * Он устанавливает значение в ячейке A2, а затем применяет к нему одинарное подчеркивание.
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
        // This example specifies that the contents of the current cell is displayed along with a line appearing directly below the character.
        
        // How to add underline to the cell value.
        
        // Get a range and add underline to its text.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A2").SetValue("The text underlined with a single line");
        worksheet.GetRange("A2").SetUnderline("single");
        worksheet.GetRange("A4").SetValue("Normal text");
        
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
