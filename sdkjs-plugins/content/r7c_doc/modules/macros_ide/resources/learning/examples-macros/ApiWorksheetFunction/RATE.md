/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.RATE
 * 
 *  Демонстрация использования метода RATE класса ApiWorksheetFunction
 * https://r7-consult.ru/
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
        // This example shows how to return the interest rate per period for a loan or an investment. For example, use 6%/4 for quarterly payments at 6% APR.
        
        // How to estimate the interest rate per period for a loan or an investment.
        
        // Use a function to get the loan/investment interest rate per period.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.RATE(2*12, -500, 10000, 0));
        
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
