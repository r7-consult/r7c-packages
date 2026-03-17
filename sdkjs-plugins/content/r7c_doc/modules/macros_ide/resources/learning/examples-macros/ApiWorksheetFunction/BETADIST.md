/**
 * @file BETADIST_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.BETADIST
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the beta probability distribution function using ApiWorksheetFunction.BETADIST.
 * It calculates the beta probability distribution function for specified parameters and displays the result in cell B2.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть функцию плотности вероятности бета-распределения с помощью ApiWorksheetFunction.BETADIST.
 * Он вычисляет функцию плотности вероятности бета-распределения для указанных параметров и отображает результат в ячейке B2.
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
        // This example shows how to return the beta probability distribution function.
        
        // How to get a result from beta probability distribution function.
        
        // Use function to get the beta probability distribution function.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let ans = func.BETADIST(0.4, 4, 5);
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
