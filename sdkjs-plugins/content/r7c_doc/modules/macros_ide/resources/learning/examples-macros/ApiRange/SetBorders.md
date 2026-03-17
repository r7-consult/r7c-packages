/**
 * @file SetBorders_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetBorders
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the border to a cell with specified parameters.
 * It sets a thick, orange bottom border to cell A2 and then sets a value in it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить границу ячейки с указанными параметрами.
 * Он устанавливает толстую оранжевую нижнюю границу для ячейки A2, а затем устанавливает в ней значение.
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
        // This example sets the border to the cell with the parameters specified.
        
        // How to set the thick bottom border to a cell.
        
        // Get a range and set its border specifying its side, type and color.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetColumnWidth(0, 50);
        worksheet.GetRange("A2").SetBorders("Bottom", "Thick", Api.CreateColorFromRGB(255, 111, 61));
        worksheet.GetRange("A2").SetValue("This is a cell with a bottom border");
        
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
