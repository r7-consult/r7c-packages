/**
 * @file COMBINA_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.COMBINA
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the number of combinations with repetitions for a given number of items using ApiWorksheetFunction.COMBINA.
 * It calculates the number of combinations with repetitions for 7 items taken 2 at a time and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть количество комбинаций с повторениями для заданного числа элементов с помощью ApiWorksheetFunction.COMBINA.
 * Он вычисляет количество комбинаций с повторениями для 7 элементов, взятых по 2 за раз, и отображает результат в ячейке A1.
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
        // This example shows how to return the number of combinations with repetitions for a given number of items.
        
        // How to find a number of combinations with repetitions.
        
        // Use function to count possible combinations with duplicates for a given number of items.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.COMBINA(7, 2));
        
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
