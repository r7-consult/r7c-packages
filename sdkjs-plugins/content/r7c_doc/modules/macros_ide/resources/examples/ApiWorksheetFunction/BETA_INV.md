/**
 * @file BETA_INV_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.BETA_INV
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the inverse of the cumulative beta probability density function for a specified beta distribution (BETADIST) using ApiWorksheetFunction.BETA_INV.
 * It calculates the inverse of the cumulative beta probability density function for specified parameters and displays the result in cell B2.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть обратную кумулятивную функцию плотности вероятности бета-распределения для заданного бета-распределения (BETADIST) с помощью ApiWorksheetFunction.BETA_INV.
 * Он вычисляет обратную кумулятивную функцию плотности вероятности бета-распределения для указанных параметров и отображает результат в ячейке B2.
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
        // This example shows how to return the inverse of the cumulative beta probability density function for a specified beta distribution (BETADIST).
        
        // How to get a result from inverse of the cumulative beta probability density function.
        
        // Use function to get the cumulative beta probability density function.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let ans = func.BETA_INV(0.2, 4, 5);
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
