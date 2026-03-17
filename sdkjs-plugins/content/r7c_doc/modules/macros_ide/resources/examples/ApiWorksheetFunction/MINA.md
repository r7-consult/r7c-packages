/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.MINA
 * 
 *  Демонстрация использования метода MINA класса ApiWorksheetFunction
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
        // This example shows how to return the smallest value in a set of values. Does not ignore logical values and text.
        
        // How to get a minimum from a list including text and logical values.
        
        // Use a function to find a minimum from a list of objects.
        
        const worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let minA = func.MINA(23, 45, true, "text", 0.89);
        worksheet.GetRange("C1").SetValue(minA);
        
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
