/**
 * @file BETAINV_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.BETAINV
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the inverse of the cumulative beta probability density function (BETA_DIST) using ApiWorksheetFunction.BETAINV.
 * It calculates the inverse of the cumulative beta probability density function for specified parameters and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть обратную кумулятивную функцию плотности вероятности бета-распределения (BETA_DIST) с помощью ApiWorksheetFunction.BETAINV.
 * Он вычисляет обратную кумулятивную функцию плотности вероятности бета-распределения для указанных параметров и отображает результат в ячейке A1.
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
        // This example shows how to return the inverse of the cumulative beta probability density function (BETA_DIST).
        
        // How to get a result from the inverse of the cumulative beta probability density function.
        
        // Use function to get the inverse of the cumulative beta probability distribution function.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.BETAINV(0.2, 4, 5));
        
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
