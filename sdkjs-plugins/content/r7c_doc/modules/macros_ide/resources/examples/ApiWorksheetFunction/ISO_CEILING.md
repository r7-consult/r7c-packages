/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.ISO_CEILING
 * 
 *  Демонстрация использования метода ISO_CEILING класса ApiWorksheetFunction
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
        // This example shows how to return a number that is rounded up to the nearest integer or to the nearest multiple of significance regardless of the sign of the number. The number is always rounded up regardless of its sing. 
        
        // How to round up a number to the nearest integer.
        
        // Use a function to round up a number to the nearest integer.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.ISO_CEILING(-6.7, 2));
        
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
