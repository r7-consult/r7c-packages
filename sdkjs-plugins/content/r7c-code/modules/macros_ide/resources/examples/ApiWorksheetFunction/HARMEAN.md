/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.HARMEAN
 * 
 *  Демонстрация использования метода HARMEAN класса ApiWorksheetFunction
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
        // This example shows how to return the harmonic mean of a data set of positive numbers: the reciprocal of the arithmetic mean of reciprocals.
        
        // How to calculate the harmonic mean of a data set of positive numbers.
        
        // Use a function to calculate harmonic mean.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let ans = func.HARMEAN(28, 16, 878, 800, 1650, 2000);
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
