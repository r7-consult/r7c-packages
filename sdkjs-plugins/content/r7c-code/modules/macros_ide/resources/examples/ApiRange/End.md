/**
 * @file End_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.End
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get a Range object that represents the end in the specified direction within a given range.
 * It gets a range and then fills the cell at the end of the range in the "xlToLeft" direction with a color.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект Range, представляющий конец в указанном направлении в заданном диапазоне.
 * Он получает диапазон, а затем заполняет ячейку в конце диапазона в направлении «xlToLeft» цветом.
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
        // This example shows how to get a Range object that represents the end in the specified direction in the specified range.
        
        // Get a left end part of a range and fill it with color.
        
        // Get a specified direction end of a range.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("C4:D5");
        range.End("xlToLeft").SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        
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
