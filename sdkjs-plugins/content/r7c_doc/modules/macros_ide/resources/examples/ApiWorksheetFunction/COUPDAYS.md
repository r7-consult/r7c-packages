/**
 * @file COUPDAYS_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.COUPDAYS
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the number of days in the coupon period that contains the settlement date.
 * It calculates the number of days in the coupon period for specified dates and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть количество дней в купонном периоде, который содержит дату расчетов.
 * Он вычисляет количество дней в купонном периоде для указанных дат и отображает результат в ячейке A1.
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
        // This example shows how to return the number of days in the coupon period that contains the settlement date.
        
        // How to find the the number of days in the coupon period that contains the settlement date.
        
        // Use function to get the number of days in the coupon period that contains the settlement date.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.COUPDAYS("1/10/2018", "6/15/2019", 4, 1));
        
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
