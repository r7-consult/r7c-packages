/**
 * @file ACCRINTM_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.ACCRINTM
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the accrued interest for a security that pays interest at maturity.
 * It calculates the accrued interest for a security with specified issue date, maturity date, annual rate, par value, and day count basis, then displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть начисленные проценты по ценной бумаге, которая выплачивает проценты при погашении.
 * Он вычисляет начисленные проценты по ценной бумаге с указанной датой выпуска, датой погашения, годовой ставкой, номинальной стоимостью и базой подсчета дней, а затем отображает результат в ячейке A1.
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
        // This example shows how to return the accrued interest for a security that pays interest at maturity.
        
        // How to get an accrued interest for a security that pays periodic interest at maturity.
        
        // Get a function that gets accrued interest for a security at maturity.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.ACCRINTM("1/1/2018", "10/15/2018", "3.50%", 1000, 1));
        
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
