/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.PPMT
 * 
 *  Демонстрация использования метода PPMT класса ApiWorksheetFunction
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
        // This example shows how to return the payment on the principal for a given investment based on periodic, constant payments and a constant interest rate.
        
        // How to return the payment on the principal for a given investment.
        
        // Use a function to calculate the payment on the principal for a given investment based on parameters.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.PPMT(0.1/12, 1, 2*12, 2000, 0));
        
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
