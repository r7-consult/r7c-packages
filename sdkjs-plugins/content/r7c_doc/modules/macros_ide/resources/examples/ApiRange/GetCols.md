/**
 * @file GetCols_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetCols
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get a Range object that represents the columns in the specified range.
 * It gets a range and then fills its second column with a color.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект Range, представляющий столбцы в указанном диапазоне.
 * Он получает диапазон, а затем заполняет его второй столбец цветом.
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
        // This example shows how to get a Range object that represents the columns in the specified range.
        
        // How to get columns from a range.
        
        // Get a range, get its first two columns and fill them with a color.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1:C3");
        range.GetCols(2).SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        
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
