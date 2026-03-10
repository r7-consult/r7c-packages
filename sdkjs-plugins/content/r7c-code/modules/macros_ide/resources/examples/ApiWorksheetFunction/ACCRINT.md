/**
 * @file ACCRINT_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.ACCRINT
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the accrued interest for a security that pays periodic interest.
 * It calculates the accrued interest for a security with specified issue date, first interest date, settlement date, annual rate, par value, and frequency, then displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть начисленные проценты по ценной бумаге, которая выплачивает периодические проценты.
 * Он вычисляет начисленные проценты по ценной бумаге с указанной датой выпуска, датой первого процента, датой расчетов, годовой ставкой, номинальной стоимостью и частотой, а затем отображает результат в ячейке A1.
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
        // This example shows how to return the accrued interest for a security that pays periodic interest.
        
        // How to get an accrued interest for a security that pays periodic interest.
        
        // Get a function that gets accrued interest for a security.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.ACCRINT("1/1/2018", "6/25/2018", "10/15/2018", "3.50%", 1000, 2));
        
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
