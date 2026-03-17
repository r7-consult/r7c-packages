/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.WEIBULL_DIST
 * 
 *  Демонстрация использования метода WEIBULL_DIST класса ApiWorksheetFunction
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
        // This example shows how to return the Weibull distribution.
        
        // How to return the Weibull distribution.
        
        // Use a function to calculate the Weibull distribution.
        
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let result = func.WEIBULL_DIST(12, 2, 5, true);
        worksheet.GetRange("B2").SetValue(result);
        
        
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
