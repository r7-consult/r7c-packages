/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.FLOOR_PRECISE
 * 
 *  Демонстрация использования метода FLOOR_PRECISE класса ApiWorksheetFunction
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
        // This example shows how to return a number that is rounded down to the nearest integer or to the nearest multiple of significance. The number is always rounded down regardless of its sign.
        
        // How to round a number down to the nearest integer regardless its sign.
        
        // Use function to round down a negative/positive number.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.FLOOR_PRECISE(-6.7, 2));
        
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
