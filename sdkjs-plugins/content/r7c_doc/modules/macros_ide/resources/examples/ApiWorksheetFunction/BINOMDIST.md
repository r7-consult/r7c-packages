/**
 * @file BINOMDIST_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.BINOMDIST
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the individual term binomial distribution probability.
 * It calculates the binomial distribution probability for specified parameters and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть вероятность биномиального распределения отдельного члена.
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
        // This example shows how to return the individual term binomial distribution probability. 
        
        // How to get an individual term binomial distribution probability.
        
        // Use function to get an individual term binomial distribution probability.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.BINOMDIST(50, 67, 0.45, false));
        
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
