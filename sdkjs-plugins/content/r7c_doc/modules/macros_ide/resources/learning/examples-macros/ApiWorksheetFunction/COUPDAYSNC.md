/**
 * @file COUPDAYSNC_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.COUPDAYSNC
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the number of days from the settlement date to the next coupon date.
 * It calculates the number of days between specified dates and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть количество дней от даты расчетов до следующей даты купона.
 * Он вычисляет количество дней между указанными датами и отображает результат в ячейке A1.
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
        // This example shows how to return the number of days from the settlement date to the next coupon date.
        
        // How to find the number of days from the settlement date to the next coupon date.
        
        // Use function to get the number of days from the settlement date to the next coupon date.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.COUPDAYSNC("1/10/2018", "6/15/2019", 4, 1));
        
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
