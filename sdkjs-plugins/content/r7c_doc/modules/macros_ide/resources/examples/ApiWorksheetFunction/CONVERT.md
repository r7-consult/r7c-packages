/**
 * @file CONVERT_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.CONVERT
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to convert a number from one measurement system to another using ApiWorksheetFunction.CONVERT.
 * It converts 2 pounds (lbm) to kilograms (kg) and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как преобразовать число из одной системы измерения в другую с помощью ApiWorksheetFunction.CONVERT.
 * Он преобразует 2 фунта (фунт) в килограммы (кг) и отображает результат в ячейке A1.
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
        // This example shows how to convert a number from one measurement system to another.
        
        // How to convert a number from one measurement system to another.
        
        // Use function to convert a number from one measurement system to another.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.CONVERT(2, "Ibm", "kg"));
        
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
