/**
 * @file Insert_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.Insert
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to insert a cell or a range of cells into the worksheet and shifts other cells away to make space.
 * It sets values in cells B4, C4, D4, and C5, and then inserts a cell at C4, shifting cells down.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вставить ячейку или диапазон ячеек в лист и сдвинуть другие ячейки, чтобы освободить место.
 * Он устанавливает значения в ячейках B4, C4, D4 и C5, а затем вставляет ячейку в C4, сдвигая ячейки вниз.
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
        // This example inserts a cell or a range of cells into the worksheet or macro sheet and shifts other cells away to make space.
        
        // How to insert a range or a cell into a worksheet.
        
        // Insert a range or a cell into a worksheet specifying its shift direction.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B4").SetValue("1");
        worksheet.GetRange("C4").SetValue("2");
        worksheet.GetRange("D4").SetValue("3");
        worksheet.GetRange("C5").SetValue("5");
        let range = worksheet.GetRange("C4");
        range.Insert("down");
        
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
