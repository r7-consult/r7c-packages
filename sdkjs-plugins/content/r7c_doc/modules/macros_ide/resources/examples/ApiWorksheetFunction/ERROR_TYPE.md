/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.ERROR_TYPE
 * 
 *  Демонстрация использования метода ERROR_TYPE класса ApiWorksheetFunction
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
        // This example shows how to return a number matching an error value.
        
        // How to get the error type code from the value.
        
        // Use function to get a error type.
        
        const worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let nonPositiveNum = 0;
        let logResult = func.LOG(nonPositiveNum);
        worksheet.GetRange("B3").SetValue(logResult);
        worksheet.GetRange("C3").SetValue(func.ERROR_TYPE(logResult));
        
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
