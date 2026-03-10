/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.NA
 * 
 *  Демонстрация использования метода NA класса ApiWorksheetFunction
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
        // This example shows how to return the N/A error value which means "no value is available".
        
        // How to return the N/A.
        
        // Use a function to get a N/A error value.
        
        const worksheet = Api.GetActiveSheet(); 
        let func = Api.GetWorksheetFunction();
        let result = func.NA();
        worksheet.GetRange("C3").SetValue(result);
        
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
