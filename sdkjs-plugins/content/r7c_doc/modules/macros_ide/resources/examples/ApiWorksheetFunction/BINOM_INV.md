/**
 * @file BINOM_INV_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.BINOM_INV
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.
 * It calculates the smallest value for a given number of trials, probability of success, and alpha value, and displays the result in cell B2.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть наименьшее значение, для которого кумулятивное биномиальное распределение больше или равно критериальному значению.
 * Он вычисляет наименьшее значение для заданного количества испытаний, вероятности успеха и значения альфа, а затем отображает результат в ячейке B2.
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
        // This example shows how to return the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value. 
        
        // How to get a smallest value for which the cumulative binomial distribution >= criterion value.
        
        // Use function to get a minimum value so that the cumulative binomial distribution >= criterion value.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let ans = func.BINOM_INV(678, 0.1, 0.007);
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
