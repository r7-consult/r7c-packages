/**
 * @file COS_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.COS
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the cosine of an angle using ApiWorksheetFunction.COS.
 * It calculates the cosine of 1.5 radians and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить косинус угла с помощью ApiWorksheetFunction.COS.
 * Он вычисляет косинус 1,5 радиан и отображает результат в ячейке A1.
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
        // This example shows how to get the cosine of an angle.
        
        // How to find a cosine.
        
        // Use function to get the cosine of an angle.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.COS(1.5));
        
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
