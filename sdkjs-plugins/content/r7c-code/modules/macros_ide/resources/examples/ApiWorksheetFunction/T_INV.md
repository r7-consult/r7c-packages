/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.T_INV
 * 
 *  Демонстрация использования метода T_INV класса ApiWorksheetFunction
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
        // This example shows how to return the left-tailed inverse of the Student's t-distribution.
        
        // How to calculate the left-tailed inverse of Student's t-distribution.
        
        // Use a function to estimate the Student's t-distribution left-tailed inverse.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let result = func.T_INV(0.75, 2);
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
