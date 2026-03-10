/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.FIXED
 * 
 *  Демонстрация использования метода FIXED класса ApiWorksheetFunction
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
        // This example shows how to round a number to the specified number of decimals and returns the result as text with or without commas.
        
        // How to round a number with parameters.
        
        // Use function to round a number with specified decimal points.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.FIXED(1234.9, 1, false));
        
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
