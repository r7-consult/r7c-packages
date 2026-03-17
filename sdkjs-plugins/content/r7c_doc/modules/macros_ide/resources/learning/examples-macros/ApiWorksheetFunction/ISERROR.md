/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.ISERROR
 * 
 *  Демонстрация использования метода ISERROR класса ApiWorksheetFunction
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
        // This example shows how to check whether a value is an error, and returns true or false.
        
        // How to check if the cell contains an error.
        
        // Use a function to check whether the value is error or not.
        
        const worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("B3").SetValue("#N/A")
        let result = func.ISERR("B3");
        worksheet.GetRange("C3").SetValue(result)
        
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
