/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.RECEIVED
 * 
 *  Демонстрация использования метода RECEIVED класса ApiWorksheetFunction
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
        // This example shows how to return the amount received at maturity for a fully invested security.
        
        // How to estimate the amount received at maturity.
        
        // Use a function to calculate the funds got at maturity for a fully invested security.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.RECEIVED("1/1/2017", "6/1/2019", "$10,000.00", "3.75%", 2));
        
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
