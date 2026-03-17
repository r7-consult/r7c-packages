/**
 * @file BETA_DIST_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.BETA_DIST
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the cumulative beta probability density function using ApiWorksheetFunction.BETA_DIST.
 * It calculates the cumulative beta probability density function for specified parameters and displays the result in cell B2.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть кумулятивную функцию плотности вероятности бета-распределения с помощью ApiWorksheetFunction.BETA_DIST.
 * Он вычисляет кумулятивную функцию плотности вероятности бета-распределения для указанных параметров и отображает результат в ячейке B2.
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
        // This example shows how to return the cumulative beta probability density function.
        
        // How to get a result from cumulative beta probability density function.
        
        // Use function to get the cumulative beta probability density function.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let ans = func.BETA_DIST(0.4, 4, 5, false);
        worksheet.GetRange("B2").SetValue(ans);
        
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
