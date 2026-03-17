/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.VAR
 * 
 *  Демонстрация использования метода VAR класса ApiWorksheetFunction
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
        // This example shows how to estimate variance based on a sample (ignores logical values and text in the sample).
        
        // How to estimate variance based on a sample.
        
        // Use a function to estimate variance ignoring logical and text values.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let result = func.VAR(1, 4, 2, 6, 0, 1, 3, 8, 0, 0, 6, 10, 0, 0, 7, 12)
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
