/**
 * @file ATAN2_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.ATAN2
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the arctangent of the specified x and y coordinates, in radians between -Pi and Pi, excluding -Pi.
 * It calculates the arctangent of coordinates (1, -9) and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть арктангенс указанных координат x и y в радианах в диапазоне от -Pi до Pi, исключая -Pi.
 * Он вычисляет арктангенс координат (1, -9) и отображает результат в ячейке A1.
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
        // This example shows how to return the arctangent of the specified x and y coordinates, in radians between -Pi and Pi, excluding -Pi.
        
        // How to get an arctangent of the specified x and y coordinates.
        
        // Use function to get an arctangent of the specified x and y coordinates in radians.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.ATAN2(1, -9));
        
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
