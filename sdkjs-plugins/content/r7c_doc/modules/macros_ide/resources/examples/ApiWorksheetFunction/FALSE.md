/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.FALSE
 * 
 *  Демонстрация использования метода FALSE класса ApiWorksheetFunction
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
        // This example shows how to return the false logical value.
        
        // How to get false value.
        
        // Use function to get a boolean false.
        
        const worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let ans = func.FALSE(); //returns false, doesnt require arguments
        worksheet.GetRange("C1").SetValue(ans);
        
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
