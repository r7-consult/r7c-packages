/**
 * @file FreezeColumns_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiFreezePanes.FreezeColumns
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to freeze a specified number of columns.
 * It gets the freeze panes object and then freezes the first column.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как закрепить указанное количество столбцов.
 * Он получает объект закрепленных областей, а затем закрепляет первый столбец.
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
        // This example freezes the the first column.
        
        // How to freeze columns using their indices.
        
        // Get freeze panes and freeze a column using its index.
        
        let worksheet = Api.GetActiveSheet();
        let freezePanes = worksheet.GetFreezePanes();
        freezePanes.FreezeColumns(1);
        
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
