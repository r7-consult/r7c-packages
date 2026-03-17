/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.MAXA
 * 
 *  Демонстрация использования метода MAXA класса ApiWorksheetFunction
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
        // This example shows how to return the largest value in a set of values. Does not ignore logical values and text.
        
        // How to get a maximum from a list including text and logical values.
        
        // Use a function to find a maximum from a list of objects.
        
        const worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1:A5").GetValue();
        let func = Api.GetWorksheetFunction();
        let maxA = func.MAX(23, 45, true, "text", 0.89);
        worksheet.GetRange("C1").SetValue(maxA);
        
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
