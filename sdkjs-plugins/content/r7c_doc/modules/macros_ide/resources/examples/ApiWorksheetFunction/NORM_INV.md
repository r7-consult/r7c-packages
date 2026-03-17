/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.NORM_INV
 * 
 *  Демонстрация использования метода NORM_INV класса ApiWorksheetFunction
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
        // This example shows how to return the inverse of the normal cumulative distribution for the specified mean and standard deviation.
        
        // How to calculate the inverse of the normal cumulative distribution.
        
        // Use a function to get the inverse of the normal cumulative distribution with a mean and standard deviation.
        
        const worksheet = Api.GetActiveSheet();
        
        worksheet.GetRange("A1").SetValue(0.6);
        
        //method params
        let value = worksheet.GetRange("A1").GetValue();
        
        let func = Api.GetWorksheetFunction();
        let ans = func.NORMSDIST(value);
        
        worksheet.GetRange("C1").SetValue(ans);
        
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
