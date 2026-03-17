/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.NORM_S_DIST
 * 
 *  Демонстрация использования метода NORM_S_DIST класса ApiWorksheetFunction
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
        // This example shows how to return the standard normal distribution (has a mean of zero and a standard deviation of one).
        
        // How to calculate the standard normal distribution.
        
        // Use a function to get the standard normal distribution with a mean = 0 and standard deviation = 1.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.NORM_S_DIST(1.33, true));
        
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
