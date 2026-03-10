/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.GESTEP
 * 
 *  Демонстрация использования метода GESTEP класса ApiWorksheetFunction
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
        // This example shows how to test whether a number is greater than a threshold value. The function returns 1 if the number is greater than or equal to the threshold value and 0 otherwise.
        
        // How to compare a number with a threshold value.
        
        // Use a function to find out whether a value greater than a limit.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.GESTEP(-2, 2));
        
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
