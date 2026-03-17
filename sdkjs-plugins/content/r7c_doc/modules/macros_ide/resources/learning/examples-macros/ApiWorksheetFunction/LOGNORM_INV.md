/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.LOGNORM_INV
 * 
 *  Демонстрация использования метода LOGNORM_INV класса ApiWorksheetFunction
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
        // This example shows how to return the inverse of the lognormal cumulative distribution function of x, where ln(x) is normally distributed with the specified parameters.
        
        // How to get the inverse of the lognormal cumulative distribution function of x.
        
        // Use a function to return the inverse of the lognormal cumulative distribution function.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.LOGNORM_INV(0.3, 2, 0.2));
        
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
