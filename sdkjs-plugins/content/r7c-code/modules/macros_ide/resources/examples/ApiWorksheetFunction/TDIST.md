/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.TDIST
 * 
 *  Демонстрация использования метода TDIST класса ApiWorksheetFunction
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
        // This example shows how to return the Student's t-distribution.
        
        // How to calculate a Student's t-distribution.
        
        // Use a function to estimate a Student's t-distribution.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let ans = func.TDIST(1.5, 10, 1);
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
