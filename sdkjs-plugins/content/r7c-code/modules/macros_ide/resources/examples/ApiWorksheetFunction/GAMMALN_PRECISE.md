/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.GAMMALN_PRECISE
 * 
 *  Демонстрация использования метода GAMMALN_PRECISE класса ApiWorksheetFunction
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
        // This example shows how to return the natural logarithm of the gamma function.
        
        // How to calculate the natural logarithm of the gamma function.
        
        // Use a function to calculate the natural logarithm of the gamma function value.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let ans = func.GAMMALN_PRECISE(0.5);
        worksheet.GetRange("B2").SetValue(ans);
        
        
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
