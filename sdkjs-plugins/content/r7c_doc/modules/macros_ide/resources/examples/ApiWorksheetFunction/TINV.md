/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.TINV
 * 
 *  Демонстрация использования метода TINV класса ApiWorksheetFunction
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
        // This example shows how to return the two-tailed inverse of the Student's t-distribution.
        
        // How to create a serial number from the two-tailed inverse.
        
        // Use a function to get two-tailed inverse of the Student's t-distribution.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let result = func.TINV(0.5, 10);
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
