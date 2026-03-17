/**
 * @file GetRows_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetRows
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get a Range object that represents the rows in the specified range.
 * It gets a range and then sets a value for each row within that range.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект Range, представляющий строки в указанном диапазоне.
 * Он получает диапазон, а затем устанавливает значение для каждой строки в этом диапазоне.
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
        // This example shows how to get a Range object that represents the rows in the specified range.
        
        // How to get a cell rows of a range.
        
        // Get a range and change each cell's row value by getting all row objects.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("1:3");
        for (let i=1; i <= 3; i++) {
        	let rows = range.GetRows(i);    
        	rows.SetValue(i);
        }
        
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
