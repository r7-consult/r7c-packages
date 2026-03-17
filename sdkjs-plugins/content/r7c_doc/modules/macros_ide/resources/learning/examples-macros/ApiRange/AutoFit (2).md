/**
 * @file AutoFit_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.AutoFit
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to auto-fit the width of columns or the height of rows in a range.
 * It sets a value in cell A1 and then auto-fits the column width to achieve the best fit.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как автоматически подогнать ширину столбцов или высоту строк в диапазоне.
 * Он устанавливает значение в ячейке A1, а затем автоматически подгоняет ширину столбца для достижения наилучшего соответствия.
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
        // This example changes the width of the columns or the height of the rows in the range to achieve the best fit.
        
        // How to set an autofit for width or height for a range.
        
        // Get a range and apply autofit property.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1");
        range.SetValue("This is an example of the column width autofit.");
        range.AutoFit(false, true);
        
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
