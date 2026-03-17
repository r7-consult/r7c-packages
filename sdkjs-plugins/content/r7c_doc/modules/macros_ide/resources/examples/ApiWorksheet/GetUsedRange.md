/**
 * @file GetUsedRange_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetUsedRange
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the ApiRange object that represents the used range on the specified worksheet.
 * It retrieves the used range and then sets its fill color.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект ApiRange, представляющий используемый диапазон на указанном листе.
 * Он извлекает используемый диапазон, а затем устанавливает его цвет заливки.
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
        // This example shows how to get the ApiRange object that represents the used range on the specified worksheet.
        
        // How to get used ranges from the worksheet.
        
        // Get used ranges and fill it with color.
        
        let worksheet = Api.GetActiveSheet();
        let usedRange = worksheet.GetUsedRange();
        usedRange.SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        
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
