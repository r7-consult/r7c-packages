/**
 * @file BINOM_DIST_RANGE_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.BINOM_DIST_RANGE
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the probability of a trial result using a binomial distribution.
 * It calculates the binomial distribution probability for specified parameters and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть вероятность результата испытания с использованием биномиального распределения.
 * Он вычисляет вероятность биномиального распределения для указанных параметров и отображает результат в ячейке A1.
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
        // This example shows how to return the probability of a trial result using a binomial distribution.
        
        // How to get a probability of a trial result.
        
        // Use function to get a probability of a trial result using binomial distribution.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.BINOM_DIST_RANGE(60, 0.75, 45, 50));
        
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
