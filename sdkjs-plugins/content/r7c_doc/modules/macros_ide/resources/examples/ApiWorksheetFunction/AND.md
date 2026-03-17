/**
 * @file AND_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.AND
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to check whether all conditions in a test are true using ApiWorksheetFunction.AND.
 * It evaluates a logical AND operation for a list of conditions and displays the result in cell C1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как проверить, истинны ли все условия в тесте, используя ApiWorksheetFunction.AND.
 * Он вычисляет логическую операцию И для списка условий и отображает результат в ячейке C1.
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
        // This example shows how to check whether all conditions in a test are true.
        
        // How to apply logical AND operation for a list of conditions.
        
        // Use logical AND to evaluate an expression.
        
        const worksheet = Api.GetActiveSheet();
        
        let func = Api.GetWorksheetFunction();
        let ans = func.AND(12 < 100, 34 < 100, 50 < 100); //AND logical function
        
        worksheet.GetRange("C1").SetValue(ans);
        
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
