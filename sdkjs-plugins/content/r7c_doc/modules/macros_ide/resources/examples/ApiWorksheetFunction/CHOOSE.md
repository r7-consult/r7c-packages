/**
 * @file CHOOSE_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.CHOOSE
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to choose a value or action to perform from a list of values, based on an index number.
 * It chooses the value at index 2 (which is 4) from the list (3, 4, 89, 76, 0) and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как выбрать значение или действие для выполнения из списка значений на основе номера индекса.
 * Он выбирает значение с индексом 2 (которое равно 4) из списка (3, 4, 89, 76, 0) и отображает результат в ячейке A1.
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
        // This example shows how to choose a value or action to perform from a list of values, based on an index number.
        
        // How to choose a value or action to perform from a list of values, based on an index number.
        
        // Use function to choose a value or action to perform from a list of values, based on an index number.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.CHOOSE(2, 3, 4, 89, 76, 0));
        
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
