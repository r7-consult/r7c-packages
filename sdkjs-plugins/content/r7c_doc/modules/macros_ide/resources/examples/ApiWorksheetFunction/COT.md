/**
 * @file COT_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.COT
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the cotangent of a number using ApiWorksheetFunction.COT.
 * It calculates the cotangent of 0.785398 radians and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить котангенс числа с помощью ApiWorksheetFunction.COT.
 * Он вычисляет котангенс 0,785398 радиан и отображает результат в ячейке A1.
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
        // This example shows how to get the cotangent of a number.
        
        // How to find a cotangent.
        
        // Use function to get the cotangent of an angle.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.COT(0.785398));
        
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
