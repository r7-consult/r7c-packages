/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.OR
 * 
 *  Демонстрация использования метода OR класса ApiWorksheetFunction
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
        // This example shows how to check whether any of the arguments are true. Returns false only if all arguments are false.
        
        // How to use OR logical operator.
        
        // Use a function to apply OR operation.
        
        const worksheet = Api.GetActiveSheet();
        
        worksheet.GetRange("A1").SetValue(12);
        
        let logical1 = worksheet.GetRange("A1") < 10;
        let logical2 = 34 < 10;
        let logical3 = 50 < 10;
        
        let func = Api.GetWorksheetFunction();
        let ans = func.OR(logical1, logical2, logical3);
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
