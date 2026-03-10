/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.PHI
 * 
 *  Демонстрация использования метода PHI класса ApiWorksheetFunction
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
        // This example shows how to return the value of the density function for a standard normal distribution.
        
        // How to return the value of the density function.
        
        // Use a function to claculate the value of the density function.
        
        const worksheet = Api.GetActiveSheet();
        
        //method params
        let number = 5;
        
        worksheet.GetRange("A1").SetValue(number);
        
        let func = Api.GetWorksheetFunction();
        let ans = func.PHI(number);
        
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
